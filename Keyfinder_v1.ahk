#Requires AutoHotkey v2.0

;==============================================================================
; APP INFO UTILITY - ENHANCED VERSION WITH STABLE RESPONSIVE DARK MODE SUPPORT
;==============================================================================
; Project: AutoHotkey Application Information and Shortcut Management Utility
; Version: 1 (Enhanced Stable Responsive Dark/Light Mode Support)
; Author: AutoHotkey Community
; Last Updated: December 2024
; License: MIT License
; 
; DESCRIPTION:
; ============
; Advanced utility for capturing Windows application information, managing 
; keyboard shortcuts, and tracking accessibility compatibility. Features 
; intelligent installation system, SQLite database with automatic file backup,
; comprehensive GUI interface with STABLE RESPONSIVE DARK/LIGHT MODE SUPPORT, 
; and extensive export capabilities.
;
; NEW FEATURES (1):
; ======================
; • Enhanced Theme Stability - Improved timing and consistency for theme changes
; • Better Error Handling - More robust fallback mechanisms for theme application
; • Reduced Flickering - Stabilized theme switching with proper delays
; • Multiple Refresh Methods - Various approaches ensure all controls update properly
; • Improved Compatibility - Better support across different Windows versions
; • Enhanced Debug Capabilities - Better troubleshooting for theme issues
; • IMPROVED TEXT READABILITY - Better contrast and darker backgrounds in dark mode
;
; PREVIOUS FEATURES (1):
; ===========================
; • Enhanced Theme Detection - Force theme checking before every dialog creation
; • Intensive Theme Monitoring - More frequent theme checks during dialog operations
; • Improved Dialog Theming - All dialogs now check theme status before creation
; • Enhanced Theme Refresh - Manual and automatic theme refresh capabilities
; • Real-time Theme Updates - Immediate theme application across all dialogs
; • Better User Experience - No visual inconsistencies during theme switches
;
; CORE FEATURES:
; ==============
; • Enhanced Stable Responsive Dark/Light Mode Interface
;   - Stable theme detection with caching and consistency checks
;   - Reduced flickering through improved timing control
;   - Multiple fallback methods for maximum compatibility
;   - Enhanced error handling for robust operation
;   - Instant adaptation to Windows theme changes (sub-second response)
;   - Real-time refresh of all open windows and dialogs
;   - Enhanced Windows API-level control theming
;   - No restart required for theme switching
;   - Consistent dark theming across all dialogs
;   - Enhanced visual appearance in both light and dark modes
;   - IMPROVED text readability with proper dark backgrounds
;
; • Intelligent Installation System
;   - Automatic version management with user data preservation
;   - Windows startup integration and desktop shortcut management
;   - Complete uninstallation with data preservation options
;
; • Database Management
;   - SQLite database with automatic file backup fallback
;   - Embedded file support for compiled versions
;   - Transaction-safe operations with rollback support
;
; • Window Information Capture
;   - Real-time window property detection
;   - Process and application identification
;   - Window class and accessibility information
;
; • Keyboard Shortcut Management
;   - Comprehensive shortcut database with categorization
;   - Bulk import/export capabilities (CSV, HTML, Text)
;   - Advanced keystroke search and filtering
;   - User-friendly shortcut capture interface
;
; • Accessibility Technology Support
;   - Screen reader compatibility tracking
;   - Known issues and workaround documentation
;   - Usage pattern analysis
;
; • Advanced GUI System with Enhanced Stable Responsive Dark Mode
;   - Stable theme detection with error handling
;   - Consistent theme monitoring during active operations
;   - Instant dark/light theme detection and switching
;   - Real-time refresh of all open windows
;   - Focus management with hotkey suspension
;   - Responsive resize handling
;   - Multi-dialog coordination
;   - Comprehensive keyboard navigation
;
; GLOBAL HOTKEYS:
; ===============
; `` & 2    : Capture currently focused window
; `` & 3    : Search applications by keystroke combination
; Escape   : Close/hide any active dialog or main window
;
; SYSTEM REQUIREMENTS:
; ===================
; • Windows 7 or later (Windows 10 1903+ recommended for best dark mode support)
; • AutoHotkey v2.0 or later
; • SQLite3.dll (included with compiled version)
; • Minimum 50MB free disk space
;
; INSTALLATION PATHS:
; ==================
; Default Install: %USERPROFILE%\Documents\AppInfoUtility\
; Settings File:   %INSTALL_DIR%\settings.ini
; Database File:   %INSTALL_DIR%\appinfo.db
; Backup Files:    %INSTALL_DIR%\Backups\
; Export Files:    %INSTALL_DIR%\Exports\
; Log Files:       %INSTALL_DIR%\Logs\
;
; ARCHITECTURE:
; =============
; The application follows a modular architecture with clear separation of 
; concerns:
; 1. Installation & Configuration Management
; 2. Database Layer (SQLite + File Backup)
; 3. Enhanced Stable Responsive Dark/Light Mode Theme Management System
; 4. GUI Framework with Focus Management
; 5. Window Capture & Analysis Engine
; 6. Shortcut Management System
; 7. Export & Import Subsystem
; 8. Accessibility Compatibility Tracker
;==============================================================================

#Warn All, Off
SendMode "Input"

;==============================================================================
; GLOBAL CONSTANTS & CONFIGURATION
;==============================================================================

/** @constant {String} Current application version */
global SCRIPT_VERSION := "1"

/** @constant {String} Application name for display and registry */
global SCRIPT_NAME := "Keyfinder"

/** @constant {Array} Required files for application operation */
global REQUIRED_FILES := ["sqlite3.dll", "appinfo.db"]

;==============================================================================
; GLOBAL STATE VARIABLES
;==============================================================================

/** @global {Boolean} Suppress non-critical messages when true */
global quietMode := false

/** @global {Boolean} Show debug information when true */
global debugMode := false

/** @global {Boolean} Indicates first-time installation */
global FIRST_RUN := false

/** @global {Boolean} Global hotkeys suspended state */
global globalHotkeysSuspended := false

/** @global {Boolean} Current Windows theme state for change detection */
global CurrentTheme := false

/** @global {Boolean} Intensive theme monitoring active state */
global intensiveThemeMonitoring := false

/** @global {String} Hotkey for capturing a window */
global captureHotkey := "`` & 2"

/** @global {String} Hotkey for searching shortcuts */
global searchHotkey := "`` & 3"

;==============================================================================
; INSTALLATION & PATH VARIABLES
;==============================================================================

/** @global {String} Main installation directory path */
global INSTALL_DIR := ""

/** @global {String} Settings configuration file path */
global SETTINGS_FILE := ""

/** @global {String} SQLite database file path */
global DB_FILE := ""

;==============================================================================
; APPLICATION DATA STRUCTURES
;==============================================================================

/** @global {Integer} Handle of last detected window */
global lastDetectedHwnd := 0

/** @global {Map} In-memory shortcuts database: ProcessName => Array<ShortcutObject> */
global shortcuts_db := Map()

/** @global {Map} Accessibility compatibility database: ProcessName => Array<CompatObject> */
global at_compat_db := Map()

/** @global {Array} Accessibility usage tracking data */
global at_usage_db := []

; Global callback management variables
global activeCallbacks := Map()
global themeCallbackInProgress := false

;==============================================================================
; GUI COMPONENT REFERENCES
;==============================================================================
; Main Application Window Components
/** @global {Object} Main application GUI window */
global MyGui := 0

/** @global {Object} Shortcuts display text area */
global ShortcutsText := 0

/** @global {Object} Category filter dropdown */
global CategoryFilter := 0

/** @global {Object} Add application to database button */
global AddAppButton := 0

/** @global {Object} Application preferences button */
global PreferencesButton := 0

/** @global {Object} Close application button */
global CloseButton := 0

/** @global {Object} Help dialog button */
global HelpButton := 0

/** @global {Object} Export shortcuts button */
global ExportButton := 0

/** @global {Object} Category filter label */
global FilterLabel := 0

/** @global {Object} Shortcuts section header */
global ShortcutsHeader := 0

;==============================================================================
; DIALOG MANAGEMENT SYSTEM
;==============================================================================

/** @global {Map} Active GUI dialogs registry: HWND => CloseCallback */
global activeGuis := Map()

/** @global {Map} Focus monitoring timers: HWND => TimerFunction */
global activeFocusTimers := Map()

;==============================================================================
; EMBEDDED RESOURCE FILES
;==============================================================================
; These files are embedded during compilation for standalone distribution
; FileInstall MUST be at global scope for compiler to detect and embed files
; Files are extracted to Windows TEMP directory (hidden from user), then copied to INSTALL_DIR
if (A_IsCompiled) {
    try {
        ; Extract to Windows temp directory - never visible to user
        tempDir := A_Temp "\Keyfinder_Temp"
        if (!DirExist(tempDir)) {
            DirCreate(tempDir)
        }
        FileInstall("sqlite3.dll", tempDir "\sqlite3.dll", 1)
        FileInstall("appinfo.db", tempDir "\appinfo.db", 0)
    } catch {
        ; Silent fail - will be handled later when we try to use the files
    }
}




/**
 * Enhanced callback manager to prevent memory leaks and race conditions
 */
class ThemeCallbackManager {
    static callbacks := Map()
    static isProcessing := false
    
    /**
     * Create and manage a callback safely
     * @param {Object} gui - GUI object to process
     * @param {Function} callbackFunc - Function to bind as callback
     * @param {Any*} params - Parameters to bind to callback
     * @return {Boolean} True if callback was created and executed successfully
     */
    static CreateAndExecute(gui, callbackFunc, params*) {
        if (!IsObject(gui) || !gui.HasOwnProp("Hwnd") || !gui.Hwnd) {
            return false
        }
        
        ; Prevent multiple simultaneous callbacks on the same GUI
        guiKey := "gui_" . gui.Hwnd
        if (this.callbacks.Has(guiKey)) {
            return false  ; Already processing this GUI
        }
        
        try {
            ; Create callback with bound parameters
            callback := CallbackCreate(callbackFunc.Bind(params*), "F", 2)
            if (!callback) {
                return false
            }
            
            ; Store callback reference
            this.callbacks[guiKey] := callback
            
            ; Execute the callback
            result := DllCall("EnumChildWindows", "Ptr", gui.Hwnd, "Ptr", callback, "Ptr", 0)
            
            ; Schedule cleanup after a short delay to ensure callback completes
            SetTimer(this.CleanupCallback.Bind(this, guiKey), -500)
            
            return true
            
        } catch Error as e {
            ; Cleanup on error
            if (this.callbacks.Has(guiKey)) {
                try {
                    CallbackFree(this.callbacks[guiKey])
                } catch {
                    ; Ignore cleanup errors
                }
                this.callbacks.Delete(guiKey)
            }
            return false
        }
    }
    
    /**
     * Safely cleanup a specific callback
     * @param {String} guiKey - Key identifying the GUI callback
     */
    static CleanupCallback(guiKey) {
        if (this.callbacks.Has(guiKey)) {
            try {
                CallbackFree(this.callbacks[guiKey])
            } catch {
                ; Ignore cleanup errors
            }
            this.callbacks.Delete(guiKey)
        }
    }
    
    /**
     * Cleanup all callbacks (used during shutdown)
     */
    static CleanupAll() {
        for guiKey, callback in this.callbacks {
            try {
                CallbackFree(callback)
            } catch {
                ; Ignore cleanup errors
            }
        }
        this.callbacks.Clear()
    }
}

/**
 * Global cleanup function for theme callbacks
 * Wrapper for ThemeCallbackManager.CleanupAll()
 * @return {Void}
 */
CleanupThemeCallbacks() {
    try {
        ThemeCallbackManager.CleanupAll()
    } catch Error as e {
        ; Silently ignore cleanup errors during shutdown
        if (debugMode) {
            ShowMessage("CleanupThemeCallbacks error: " . e.Message)
        }
    }
}

;==============================================================================
; GLOBAL HOTKEY DEFINITIONS
;==============================================================================

/**
 * Dynamically registers the global hotkeys based on the current settings.
 * This function will deregister old hotkeys before registering new ones.
 */
RegisterGlobalHotkeys() {
    static registeredCaptureHotkey := ""
    static registeredSearchHotkey := ""

    ; Deregister previous hotkeys if they exist, ignoring errors if they don't.
    try {
        if (registeredCaptureHotkey != "") {
            try {
                Hotkey(registeredCaptureHotkey, "Off")
            } catch {
                ; Ignore if hotkey doesn't exist
            }
        }
    } catch {
        ; Ignore error if hotkey was invalid or couldn't be unregistered
    }
    try {
        if (registeredSearchHotkey != "") {
            try {
                Hotkey(registeredSearchHotkey, "Off")
            } catch {
                ; Ignore if hotkey doesn't exist
            }
        }
    } catch {
        ; Ignore error
    }

    ; Convert human-readable format to AutoHotkey format
    ahkCaptureHotkey := ConvertToAutoHotkeyFormat(captureHotkey)
    ahkSearchHotkey := ConvertToAutoHotkeyFormat(searchHotkey)

    ; Register new hotkeys with converted format
    try {
        if (ahkCaptureHotkey != "") {
            Hotkey(ahkCaptureHotkey, (*) => CaptureFocusedWindow())
            registeredCaptureHotkey := ahkCaptureHotkey
        }
    } catch Error as e {
        MsgBox("Failed to register the Capture Window hotkey: " . captureHotkey . "`nAutoHotkey format: " . ahkCaptureHotkey . "`nError: " . e.Message, "Hotkey Error", 16)
        registeredCaptureHotkey := ""
    }

    try {
        if (ahkSearchHotkey != "") {
            Hotkey(ahkSearchHotkey, (*) => ShowSearchShortcutDialog())
            registeredSearchHotkey := ahkSearchHotkey
        }
    } catch Error as e {
        MsgBox("Failed to register the Search Shortcut hotkey: " . searchHotkey . "`nAutoHotkey format: " . ahkSearchHotkey . "`nError: " . e.Message, "Hotkey Error", 16)
        registeredSearchHotkey := ""
    }
    
    ; Ensure other hotkeys are still active using the direct definition syntax
    try {
; Define conditional Escape key binding only when AppInfoUtilityWindows are active
GroupAdd("AppInfoUtilityWindows", "ahk_class AutoHotkeyGUI")  ; Define group once
HotIf WinActive("ahk_group AppInfoUtilityWindows")
Hotkey("Escape", (*) => CloseActiveDialogOrMainWindow())
HotIf  ; Reset to global context
    } catch {
        ; Some hotkeys might conflict, continue anyway
    }
}



; Backtick prefix hotkey (`` & key combinations)
SC029::SC029

; Enhanced hotkeys with theme checking for core functionality
SC029 & 2::{
    ; Check theme before capturing to ensure proper display
    CaptureFocusedWindow()
}

SC029 & 3::{
    ; Check theme before showing search dialog
    ForceThemeCheck()
    ShowSearchShortcutDialog()
}


;==============================================================================
; ENHANCED STABLE RESPONSIVE DARK/LIGHT MODE THEME MANAGEMENT SYSTEM
;==============================================================================


/**
 * Test hotkey registration to ensure they're working after theme changes
 * @return {Void}
 */
TestHotkeyRegistration() {
    try {
        ; Test if our main hotkeys are still registered
        ; This is called after theme changes to ensure hotkeys weren't disrupted
        
        ; Check if we can register a test hotkey
        testHotkey := "F24"  ; F24 is rarely used
        
        try {
            ; Try to register and immediately unregister a test hotkey
            Hotkey(testHotkey, (*) => "", "On")
            Hotkey(testHotkey, "Off")
            
            ; If we get here, hotkey system is working
            if (debugMode) {
                ShowMessage("Hotkey system test passed - all hotkeys should be functional")
            }
        } catch Error as e {
            ; Hotkey system might have issues, try to re-register our main hotkeys
            if (debugMode) {
                ShowMessage("Hotkey system test failed, attempting to re-register: " . e.Message)
            }
            
            ; Re-register our global hotkeys
            RegisterGlobalHotkeys()
        }
        
    } catch Error as e {
        ; Complete failure - log but don't crash
        if (debugMode) {
            ShowMessage("TestHotkeyRegistration failed: " . e.Message)
        }
    }
}


/**
 * Improved theme detection with better error handling and consistency
 * @return {Boolean} True if Windows is in dark mode, false for light mode
 */
IsWindowsDarkMode() {
    static lastResult := false
    static lastCheck := 0
    
    ; Cache result for 100ms to prevent rapid repeated calls
    currentTime := A_TickCount
    if (currentTime - lastCheck < 100) {
        return lastResult
    }
    
    try {
        ; Primary method - check AppsUseLightTheme
        lightTheme := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme")
        result := (lightTheme == 0)
        
        ; Validation - also check SystemUsesLightTheme for consistency
        try {
            systemLightTheme := RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "SystemUsesLightTheme")
            ; If there's a mismatch, prefer the app setting but log it
            if ((systemLightTheme == 0) != result) {
                ; Mixed theme mode - prefer app setting
            }
        } catch {
            ; SystemUsesLightTheme key doesn't exist, continue with app setting
        }
        
        lastResult := result
        lastCheck := currentTime
        return result
        
    } catch {
        ; Registry key doesn't exist or other error - return cached result or default
        lastCheck := currentTime
        return lastResult  ; Keep previous result if available
    }
}

/**
 * Enhanced theme colors with better contrast and darker backgrounds for text readability
 * @return {Map} Color map with improved dark mode definitions
 */
GetThemeColors() {
    isDark := IsWindowsDarkMode()
    
    if (isDark) {
        ; Enhanced dark mode colors with darker backgrounds for better text readability
        return Map(
            "background", 0x1E1E1E,        ; Darker main background (VS Code dark)
            "textColor", 0xFFFFFF,         ; Pure white text for maximum contrast
            "editBackground", 0x252526,    ; Dark edit field background
            "editText", 0xFFFFFF,          ; White edit field text
            "buttonBackground", 0x2D2D30,  ; Dark button background
            "buttonText", 0xFFFFFF,        ; White button text
            "menuBackground", 0x252526,    ; Dark menu background
            "menuText", 0xFFFFFF,          ; White menu text
            "menuHighlight", 0x094771,     ; Darker blue selection highlight
            "borderColor", 0x464647,       ; Subtle border color
            "groupBoxBackground", 0x1E1E1E, ; Dark GroupBox background
            "listViewBackground", 0x252526, ; Dark ListView background
            "scrollBackground", 0x2D2D30,  ; Dark scrollbar background
            "controlText", 0xFFFFFF,       ; White control text (fixed from dark gray)
            "secondaryBackground", 0x2D2D30, ; Secondary dark background
            "inputBackground", 0x3C3C3C,   ; Input field dark background
            "panelBackground", 0x252526    ; Panel/container dark background
        )
    } else {
        ; Enhanced light mode colors with better contrast
        return Map(
            "background", 0xFFFFFF,        ; Pure white background
            "textColor", 0x000000,         ; Pure black text
            "editBackground", 0xFFFFFF,    ; Pure white edit backgrounds
            "editText", 0x000000,          ; Black edit text
            "buttonBackground", 0xF0F0F0,  ; Light gray buttons
            "buttonText", 0x000000,        ; Black button text
            "menuBackground", 0xFFFFFF,    ; White menu background
            "menuText", 0x000000,          ; Black menu text
            "menuHighlight", 0x0078D4,     ; Blue selection highlight
            "borderColor", 0xD0D0D0,       ; Light gray borders
            "groupBoxBackground", 0xFAFAFA, ; Very light gray GroupBox
            "listViewBackground", 0xFFFFFF, ; White ListView
            "scrollBackground", 0xF0F0F0,  ; Light scrollbar
            "controlText", 0x000000,       ; Black control text
            "secondaryBackground", 0xF8F8F8, ; Secondary light background
            "inputBackground", 0xFFFFFF,   ; Input field white background
            "panelBackground", 0xFAFAFA    ; Panel/container light background
        )
    }
}

/**
 * Enhanced Windows API dark mode support with better error handling
 * @return {Void}
 */
EnableDarkModeForApp() {
    static isEnabled := false
    static lastTheme := -1
    
    try {
        isDark := IsWindowsDarkMode()
        
        ; Only update if theme actually changed
        if (lastTheme == isDark && isEnabled) {
            return
        }
        
        ; Get Windows version for compatibility
        winVer := VerCompare(A_OSVersion, "10.0.17763")  ; Windows 10 1809
        if (winVer < 0) {
            return  ; Not supported on older versions
        }
        
        if (isDark) {
            ; Enable dark mode
            try {
                ; Windows 10 1903+ dark mode attribute
                if (VerCompare(A_OSVersion, "10.0.18362") >= 0) {
                    DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", A_ScriptHwnd, "UInt", 20, "Int*", 1, "UInt", 4)
                }
                
                ; Try to set preferred app mode
                try {
                    ; Using ordinal 135 for SetPreferredAppMode (more reliable)
                    DllCall("uxtheme\#135", "Int", 1) ; APPMODE_ALLOWDARK
                    DllCall("uxtheme\FlushMenuThemes")
                } catch {
                    ; Fallback for older versions
                    try {
                        DllCall("uxtheme\SetPreferredAppMode", "Int", 1)
                        DllCall("uxtheme\FlushMenuThemes")
                    } catch {
                        ; API not available
                    }
                }
                
                ; Set window theme
                try {
                    DllCall("uxtheme\SetWindowTheme", "Ptr", A_ScriptHwnd, "WStr", "DarkMode_Explorer", "Ptr", 0)
                } catch {
                    ; Theme setting failed
                }
                
            } catch Error as e {
                ; Continue with basic theming even if advanced features fail
            }
        } else {
            ; Light mode
            try {
                if (VerCompare(A_OSVersion, "10.0.18362") >= 0) {
                    DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", A_ScriptHwnd, "UInt", 20, "Int*", 0, "UInt", 4)
                }
                
                try {
                    DllCall("uxtheme\#135", "Int", 0) ; APPMODE_DEFAULT
                    DllCall("uxtheme\FlushMenuThemes")
                } catch {
                    try {
                        DllCall("uxtheme\SetPreferredAppMode", "Int", 0)
                        DllCall("uxtheme\FlushMenuThemes")
                    } catch {
                        ; API not available
                    }
                }
                
                try {
                    DllCall("uxtheme\SetWindowTheme", "Ptr", A_ScriptHwnd, "Ptr", 0, "Ptr", 0)
                } catch {
                    ; Theme clearing failed
                }
                
            } catch Error as e {
                ; Continue with basic theming
            }
        }
        
        lastTheme := isDark
        isEnabled := true
        
        ; Force a complete refresh of all windows
        try {
            ; Send theme change message to all windows
            DllCall("user32\PostMessage", "Ptr", 0xFFFF, "UInt", 0x001A, "Ptr", 0, "Ptr", 0) ; HWND_BROADCAST, WM_WININICHANGE
        } catch {
            ; Message sending failed
        }
        
    } catch Error as e {
        ; Complete failure - continue without advanced theming
    }
}

/**
 * Enhanced ForceThemeCheck with better control refresh
 */
ForceThemeCheck() {
    global CurrentTheme
    static isProcessing := false
    
    ; Prevent multiple simultaneous theme checks
    if (isProcessing) {
        return false
    }
    
    isProcessing := true
    
    try {
        ; Store current theme state
        previousTheme := CurrentTheme
        
        ; Force detection with multiple attempts
        newTheme := false
        attempts := 0
        maxAttempts := 3
        
        while (attempts < maxAttempts) {
            try {
                newTheme := IsWindowsDarkMode()
                break
            } catch {
                attempts++
                Sleep(50)
            }
        }
        
        ; Update current theme
        CurrentTheme := newTheme
        
        ; If theme changed or this is a forced check, refresh everything
        if (CurrentTheme != previousTheme || attempts > 0) {
            ; Disable all window updates temporarily
            DllCall("LockWindowUpdate", "Ptr", DllCall("GetDesktopWindow"))
            
            try {
                RefreshApplicationTheme()
                Sleep(100) ; Allow Windows API calls to complete
                RefreshAllOpenGUIs()
                Sleep(100) ; Allow GUI updates to complete
                
                ; ADDED: Force refresh of main window specifically
                RefreshMainWindowCompletely()
                
            } finally {
                ; Re-enable window updates
                DllCall("LockWindowUpdate", "Ptr", 0)
            }
            
            return true
        }
        
        return false
        
    } finally {
        isProcessing := false
    }
}

/**
 * Complete refresh of main window with all controls
 */
RefreshMainWindowCompletely() {
    global MyGui, ShortcutsText, CategoryFilter
    global AddAppButton, PreferencesButton, CloseButton, HelpButton, ExportButton
    global FilterLabel, ShortcutsHeader
    
    try {
        if (!MyGui || !MyGui.Hwnd || !WinExist("ahk_id " MyGui.Hwnd)) {
            return
        }
        
        ; Get current theme colors
        colors := GetThemeColors()
        isDark := IsWindowsDarkMode()
        
        ; Apply theme to main window first
        ApplyDarkModeToGUI(MyGui, "Main Application Window")
        
        ; Force background color update
        MyGui.BackColor := colors["background"]
        
        ; Update each control individually with error handling
        controlList := [
            {obj: ShortcutsHeader, type: "Text"},
            {obj: ShortcutsText, type: "Edit"},
            {obj: FilterLabel, type: "Text"},
            {obj: CategoryFilter, type: "DropDownList"},
            {obj: AddAppButton, type: "Button"},
            {obj: ExportButton, type: "Button"},
            {obj: PreferencesButton, type: "Button"},
            {obj: CloseButton, type: "Button"},
            {obj: HelpButton, type: "Button"}
        ]
        
        for controlRef in controlList {
            try {
                if (IsObject(controlRef.obj) && controlRef.obj.Hwnd) {
                    ApplyControlTheming(controlRef.obj, controlRef.type, colors, isDark)
                    
                    ; Force individual control refresh
                    DllCall("InvalidateRect", "Ptr", controlRef.obj.Hwnd, "Ptr", 0, "Int", 1)
                    DllCall("UpdateWindow", "Ptr", controlRef.obj.Hwnd)
                }
            } catch {
                ; Continue with next control if this one fails
            }
        }
        
        ; Force complete window refresh with multiple methods
        try {
            MyGui.Redraw()
            DllCall("InvalidateRect", "Ptr", MyGui.Hwnd, "Ptr", 0, "Int", 1)
            DllCall("UpdateWindow", "Ptr", MyGui.Hwnd)
            DllCall("RedrawWindow", "Ptr", MyGui.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x0001 | 0x0004 | 0x0010)
        } catch {
            ; Fallback to basic redraw
            try {
                MyGui.Redraw()
            } catch {
                ; Even basic redraw failed
            }
        }
        
        ; Apply enhanced text control theming
        ForceTextControlsRetheming(MyGui)
        
        ; Schedule additional refresh for stubborn controls
        SetTimer(FinalMainWindowRefresh, -300)
        
    } catch Error as e {
        ; Complete refresh failed, but don't crash
    }
}

/**
 * Final refresh attempt for any remaining theme issues
 */
FinalMainWindowRefresh() {
    try {
        global MyGui
        
        if (!MyGui || !MyGui.Hwnd || !WinExist("ahk_id " MyGui.Hwnd)) {
            return
        }
        
        ; One more complete refresh
        colors := GetThemeColors()
        MyGui.BackColor := colors["background"]
        
        ; Force complete repaint
        DllCall("RedrawWindow", "Ptr", MyGui.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x0001 | 0x0004 | 0x0010 | 0x0100)
        
    } catch {
        ; Final refresh failed, but that's ok
    }
}

/**
 * Enhanced GUI theming with better background color application
 */
ApplyDarkModeToGUI(gui, title := "") {
    if (!IsObject(gui) || !gui.Hwnd) {
        return
    }
    
    try {
        colors := GetThemeColors()
        isDark := IsWindowsDarkMode()
        
        ; Set GUI background color with enhanced dark color
        gui.BackColor := colors["background"]
        
        ; Apply window-level dark mode with better error handling
        if (isDark) {
            try {
                ; Windows 11/10 dark mode attribute
                if (VerCompare(A_OSVersion, "10.0.22000") >= 0) {
                    DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", gui.Hwnd, "UInt", 20, "Int*", 1, "UInt", 4)
                } else if (VerCompare(A_OSVersion, "10.0.18362") >= 0) {
                    DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", gui.Hwnd, "UInt", 20, "Int*", 1, "UInt", 4)
                }
                
                ; Set window theme for better dark mode support
                DllCall("uxtheme\SetWindowTheme", "Ptr", gui.Hwnd, "WStr", "DarkMode_Explorer", "Ptr", 0)
            } catch {
                ; Windows API calls failed, continue with color-only theming
            }
        } else {
            try {
                if (VerCompare(A_OSVersion, "10.0.18362") >= 0) {
                    DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", gui.Hwnd, "UInt", 20, "Int*", 0, "UInt", 4)
                }
                DllCall("uxtheme\SetWindowTheme", "Ptr", gui.Hwnd, "Ptr", 0, "Ptr", 0)
            } catch {
                ; API calls failed
            }
        }
        
        ; Set default font with appropriate color and ensure it applies to all text
        try {
            fontColor := Format("c0x{:06X}", colors["textColor"])
            gui.SetFont("s10 " fontColor)
        } catch {
            ; Font setting failed
        }
        
        ; Force complete window refresh with multiple methods
        try {
            gui.Redraw()
            
            ; Additional refresh methods for better reliability
            DllCall("InvalidateRect", "Ptr", gui.Hwnd, "Ptr", 0, "Int", 1)
            DllCall("UpdateWindow", "Ptr", gui.Hwnd)
            
            ; Send theme change messages to ensure all child controls update
            DllCall("SendMessage", "Ptr", gui.Hwnd, "UInt", 0x001A, "Ptr", 0, "Ptr", 0) ; WM_WININICHANGE
            DllCall("SendMessage", "Ptr", gui.Hwnd, "UInt", 0x031A, "Ptr", 0, "Ptr", 0) ; WM_THEMECHANGED
            
        } catch {
            ; Refresh failed
        }
        
    } catch Error as e {
        ; Complete theming failure - continue without theming
    }
}

/**
 * Enhanced ApplyControlTheming function with better Text control handling
 */
ApplyControlTheming(control, controlType, colors, isDark) {
    if (!IsObject(control) || !IsObject(control.Gui) || !control.Hwnd)
        return

    try {
        if (isDark) {
            ; --- DARK MODE --- 
            switch controlType {
                case "Edit":
                    {
                        ; This now handles BOTH normal and ReadOnly Edit controls 
                        control.Opt("+Background" . Format("0x{:06X}", colors["editBackground"]))
                        control.SetFont("c" . Format("0x{:06X}", colors["editText"]))
                        DllCall("uxtheme\SetWindowTheme", "Ptr", control.Hwnd, "WStr", "DarkMode_CFD", "Ptr", 0)
                    }
                case "DropDownList", "ComboBox":
                    {
                        control.Opt("+Background" . Format("0x{:06X}", colors["menuBackground"]))
                        control.SetFont("c" . Format("0x{:06X}", colors["menuText"]))
                        DllCall("uxtheme\SetWindowTheme", "Ptr", control.Hwnd, "WStr", "DarkMode_CFD", "Ptr", 0)
                    }
                case "Button", "Checkbox", "Radio":
                    {
                        control.Opt("+Background" . Format("0x{:06X}", colors["buttonBackground"]))
                        control.SetFont("c" . Format("0x{:06X}", colors["buttonText"]))
                    }
                case "Text":
                    {
                        ; ENHANCED TEXT CONTROL THEMING
                        ; Clear any existing background first
                        control.Opt("-Background")
                        control.Opt("+BackgroundTransparent")
                        
                        ; Apply font color with multiple methods for reliability
                        fontColor := Format("0x{:06X}", colors["textColor"])
                        control.SetFont("c" . fontColor)
                        
                        ; Force text color using Windows API for stubborn controls
                        try {
                            ; Send WM_CTLCOLORSTATIC message to force color update
                            DllCall("InvalidateRect", "Ptr", control.Hwnd, "Ptr", 0, "Int", 1)
                            
                            ; Alternative method: Set text color directly via API
                            hdc := DllCall("GetDC", "Ptr", control.Hwnd, "Ptr")
                            if (hdc) {
                                DllCall("SetTextColor", "Ptr", hdc, "UInt", colors["textColor"])
                                DllCall("SetBkMode", "Ptr", hdc, "Int", 1) ; TRANSPARENT
                                DllCall("ReleaseDC", "Ptr", control.Hwnd, "Ptr", hdc)
                            }
                        } catch {
                            ; API calls failed, continue with basic method
                        }
                    }
                case "ListView":
                    {
                        control.Opt("+Background" . Format("0x{:06X}", colors["listViewBackground"]))
                        DllCall("uxtheme\SetWindowTheme", "Ptr", control.Hwnd, "WStr", "DarkMode_Explorer", "Ptr", 0)
                    }
                case "GroupBox":
                    {
                        control.Opt("+BackgroundTransparent")
                        control.SetFont("c" . Format("0x{:06X}", colors["textColor"]))
                    }
            }
        } else {
            ; --- LIGHT MODE --- 
            DllCall("uxtheme\SetWindowTheme", "Ptr", control.Hwnd, "Ptr", 0, "Ptr", 0)
            
            if (controlType = "Text") {
                ; Enhanced light mode for Text controls
                control.Opt("-Background")
                control.Opt("+BackgroundDefault")
                fontColor := Format("0x{:06X}", colors["textColor"])
                control.SetFont("c" . fontColor)
                
                ; Force refresh for light mode text
                try {
                    DllCall("InvalidateRect", "Ptr", control.Hwnd, "Ptr", 0, "Int", 1)
                    hdc := DllCall("GetDC", "Ptr", control.Hwnd, "Ptr")
                    if (hdc) {
                        DllCall("SetTextColor", "Ptr", hdc, "UInt", colors["textColor"])
                        DllCall("SetBkMode", "Ptr", hdc, "Int", 2) ; OPAQUE for light mode
                        DllCall("ReleaseDC", "Ptr", control.Hwnd, "Ptr", hdc)
                    }
                } catch {
                    ; API calls failed
                }
            } else {
                control.Opt("-Background") ; Revert to default background 
            }
            control.SetFont("c" . Format("0x{:06X}", colors["textColor"]))
        }
        
        ; Enhanced refresh for Text controls specifically
        if (controlType = "Text") {
            ForceTextControlRefresh(control)
        } else {
            RefreshSingleControl(control)
        }
    } catch {
        ; Silently handle any theming errors 
    }
}

/**
 * Specialized refresh function for Text controls
 */
ForceTextControlRefresh(control) {
    try {
        ; Multiple refresh methods for Text controls
        DllCall("InvalidateRect", "Ptr", control.Hwnd, "Ptr", 0, "Int", 1)
        DllCall("UpdateWindow", "Ptr", control.Hwnd)
        
        ; Force complete redraw
        DllCall("SendMessage", "Ptr", control.Hwnd, "UInt", 0x000F, "Ptr", 0, "Ptr", 0) ; WM_PAINT
        
        ; Alternative refresh method for stubborn text controls
        try {
            ; Hide and show to force complete refresh
            DllCall("ShowWindow", "Ptr", control.Hwnd, "Int", 0) ; SW_HIDE
            DllCall("ShowWindow", "Ptr", control.Hwnd, "Int", 5) ; SW_SHOW
        } catch {
            ; Hide/show failed, continue
        }
        
        ; Final force refresh
        DllCall("RedrawWindow", "Ptr", control.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x0001 | 0x0004 | 0x0010)
        
    } catch {
        ; Text control refresh failed
    }
}


/**
 * Applies a robust theme to ReadOnly Edit controls by directly handling
 * the WM_CTLCOLORSTATIC message to force the background color.
 * @param {Object} control The ReadOnly Edit control object.
 * @param {Map} colors The current theme's color map.
 * @param {Boolean} isDark True if dark mode is active.
 */

/**
 * Applies a robust theme to ReadOnly Edit controls by registering their
 * HWND and desired background color for the global WM_CTLCOLORSTATIC handler.
 * @param {Object} control The ReadOnly Edit control object.
 * @param {Map} colors The current theme's color map.
 * @param {Boolean} isDark True if dark mode is active.
 */

/**
 * Enhanced single control refresh with multiple methods
 */
RefreshSingleControl(control) {
    try {
        ; Multiple refresh methods for maximum compatibility
        DllCall("InvalidateRect", "Ptr", control.Hwnd, "Ptr", 0, "Int", 1)
        DllCall("UpdateWindow", "Ptr", control.Hwnd)
        
        ; Send specific refresh messages
        DllCall("SendMessage", "Ptr", control.Hwnd, "UInt", 0x000F, "Ptr", 0, "Ptr", 0) ; WM_PAINT
        
        ; Force redraw for certain control types
        try {
            DllCall("SendMessage", "Ptr", control.Hwnd, "UInt", 0x000B, "Ptr", 0, "Ptr", 0) ; WM_SETREDRAW FALSE
            DllCall("SendMessage", "Ptr", control.Hwnd, "UInt", 0x000B, "Ptr", 1, "Ptr", 0) ; WM_SETREDRAW TRUE
        } catch {
            ; Redraw messages failed
        }
        
    } catch {
        ; Control refresh failed
    }
}

/**
 * Enhanced ForceTextControlsRetheming with proper callback management and validation
 * @param {Object} gui - GUI object to process
 */
ForceTextControlsRetheming(gui) {
    ; Enhanced validation to prevent crashes
    if (!IsObject(gui)) {
        return
    }
    
    try {
        if (!gui.HasOwnProp("Hwnd") || !gui.Hwnd) {
            return
        }
    } catch {
        return
    }
    
    ; Verify window still exists and is valid
    try {
        if (!WinExist("ahk_id " gui.Hwnd) || !DllCall("IsWindow", "Ptr", gui.Hwnd)) {
            return
        }
    } catch {
        return
    }
    
    ; Prevent multiple simultaneous operations on the same GUI
    guiKey := "theming_" . gui.Hwnd
    global themeCallbackInProgress
    if (themeCallbackInProgress) {
        return
    }
    
    themeCallbackInProgress := true
    
    try {
        colors := GetThemeColors()
        isDark := IsWindowsDarkMode()
        
        ; Use the new callback manager instead of direct callback creation
        success := ThemeCallbackManager.CreateAndExecute(gui, ForceTextControlCallbackSafe, colors, isDark)
        
        if (!success) {
            ; Fallback to refreshing known global controls if callback method fails
            try {
                RefreshKnownGlobalControls(colors, isDark)
            } catch {
                ; All methods failed, continue silently
            }
        }
        
    } catch Error as e {
        ; Complete function failed, continue silently
    } finally {
        ; Always reset the processing flag
		SetTimer(ResetThemeCallbackFlag, -100)

		; And add this helper function:
		ResetThemeCallbackFlag() {
			global themeCallbackInProgress
			themeCallbackInProgress := false
		}
    }
}

/**
 * Enhanced and safer callback for re-theming text controls with improved validation
 * @param {Map} colors - Theme color map
 * @param {Boolean} isDark - Dark mode flag
 * @param {Integer} hwnd - Window handle to process
 * @param {Integer} lParam - Additional parameter (unused)
 * @return {Boolean} True to continue enumeration
 */
ForceTextControlCallbackSafe(colors, isDark, hwnd, lParam) {
    ; Enhanced validation to prevent crashes
    try {
        ; Verify the window handle is still valid
        if (!hwnd || !DllCall("IsWindow", "Ptr", hwnd)) {
            return true
        }
        
        ; Additional safety check - ensure window is still accessible
        try {
            ; Test if we can get basic window info
            DllCall("GetWindowThreadProcessId", "Ptr", hwnd, "Ptr*", &processId := 0, "UInt")
            if (!processId) {
                return true  ; Window is not accessible
            }
        } catch {
            return true  ; Window access failed, skip it
        }
        
        ; Get class name with enhanced error handling
        className := ""
        try {
            classBuffer := Buffer(256, 0)
            result := DllCall("GetClassName", "Ptr", hwnd, "Ptr", classBuffer, "Int", 256, "Int")
            if (result > 0) {
                className := StrGet(classBuffer, "UTF-8")
            }
        } catch {
            return true  ; Continue enumeration even if this control fails
        }
        
        ; Focus on Static (text) controls that might need better theming
        if (InStr(className, "Static")) {
            try {
                ; Apply enhanced dark mode theming with validation
                if (isDark && IsObject(colors) && colors.Has("background") && colors.Has("textColor")) {
                    ; Validate color values before using them
                    bgColor := colors["background"]
                    textColor := colors["textColor"]
                    
                    ; Only proceed if colors are valid numbers
                    if (IsInteger(bgColor) && IsInteger(textColor)) {
                        ; Use Windows API to refresh the control with error handling
                        try {
                            ; Verify window is still valid before each API call
                            if (DllCall("IsWindow", "Ptr", hwnd)) {
                                ; Invalidate and update the control
                                DllCall("InvalidateRect", "Ptr", hwnd, "Ptr", 0, "Int", 1)
                                
                                ; Only call UpdateWindow if InvalidateRect succeeded
                                DllCall("UpdateWindow", "Ptr", hwnd)
                                
                                ; Send a paint message to force immediate redraw
                                DllCall("SendMessage", "Ptr", hwnd, "UInt", 0x000F, "Ptr", 0, "Ptr", 0) ; WM_PAINT
                            }
                        } catch {
                            ; API calls failed for this control, continue with next
                        }
                    }
                }
            } catch {
                ; Control processing failed, continue with next
            }
        }
        
    } catch {
        ; Complete control processing failed, continue enumeration
    }
    
    return true  ; Continue enumeration
}


/**
 * Enhanced theme change detection with stability improvements
 * @return {Boolean} True if theme has changed
 */
CheckThemeChange() {
    global CurrentTheme
    static lastCheck := 0
    static consecutiveChanges := 0
    
    currentTime := A_TickCount
    
    ; Prevent too frequent checks (minimum 200ms between checks)
    if (currentTime - lastCheck < 200) {
        return false
    }
    
    newTheme := IsWindowsDarkMode()
    
    if (newTheme != CurrentTheme) {
        consecutiveChanges++
        
        ; Require at least 2 consecutive detections of the same change
        ; This prevents flickering from rapid theme switches
        if (consecutiveChanges >= 2) {
            CurrentTheme := newTheme
            consecutiveChanges := 0
            lastCheck := currentTime
            
            ; Perform theme refresh with delay to allow Windows to complete the switch
            SetTimer(() => DelayedThemeRefresh(), -300)
			SetTimer(() => TestHotkeyRegistration(), -2000)
            return true
        }
    } else {
        consecutiveChanges := 0
    }
    
    lastCheck := currentTime
    return false
}

/**
 * Delayed theme refresh to prevent timing issues
 */
DelayedThemeRefresh() {
    RefreshApplicationTheme()
    RefreshAllOpenGUIs()
}

/**
 * Enhanced application theme refresh
 */
RefreshApplicationTheme() {
    EnableDarkModeForApp()
    
    try {
        ; Force complete application refresh
        DllCall("uxtheme\FlushMenuThemes")
        
        ; Send theme change messages
        DllCall("user32\PostMessage", "Ptr", A_ScriptHwnd, "UInt", 0x001A, "Ptr", 0, "Ptr", 0) ; WM_WININICHANGE
        DllCall("user32\PostMessage", "Ptr", A_ScriptHwnd, "UInt", 0x031A, "Ptr", 0, "Ptr", 0) ; WM_THEMECHANGED
        
        ; Small delay to allow processing
        Sleep(50)
        
    } catch {
        ; Message sending failed
    }
}

/**
 * Enhanced RefreshAllOpenGUIs with better error handling and race condition prevention
 */
RefreshAllOpenGUIs() {
    global MyGui, activeGuis
    static isRefreshing := false
    static lastRefresh := 0
    
    ; Prevent recursive calls and too frequent refreshes
    if (isRefreshing) {
        return
    }
    
    ; Throttle refresh calls to prevent overwhelming the system
    currentTime := A_TickCount
    if (currentTime - lastRefresh < 200) {  ; Minimum 200ms between refreshes
        return
    }
    
    isRefreshing := true
    lastRefresh := currentTime
    
    try {
        ; Refresh main window if it exists with enhanced validation
        try {
            if (IsObject(MyGui)) {
                ; Multiple validation checks
                if (MyGui.HasOwnProp("Hwnd") && MyGui.Hwnd) {
                    if (WinExist("ahk_id " MyGui.Hwnd) && DllCall("IsWindow", "Ptr", MyGui.Hwnd)) {
                        RefreshMainGUITheme()
                        Sleep(50) ; Allow processing time
                    }
                }
            }
        } catch Error as e {
            ; Main window refresh failed, continue with dialogs
        }
        
        ; Refresh all active dialog windows with enhanced validation
        if (IsObject(activeGuis) && activeGuis.Count > 0) {
            ; Create a copy of the activeGuis to avoid modification during iteration
            guiList := []
            try {
                for guiHwnd, closeCallback in activeGuis {
                    guiList.Push({hwnd: guiHwnd, callback: closeCallback})
                }
            } catch {
                ; Failed to enumerate active GUIs
            }
            
            ; Process each GUI safely
            for guiInfo in guiList {
                try {
                    guiHwnd := guiInfo.hwnd
                    
                    ; Validate window still exists
                    if (guiHwnd && DllCall("IsWindow", "Ptr", guiHwnd) && WinExist("ahk_id " guiHwnd)) {
                        guiObj := GuiFromHwnd(guiHwnd)
                        if (guiObj && IsObject(guiObj)) {
                            RefreshGUITheme(guiObj, "Dialog Window")
                        } else {
                            ; Try direct refresh using Windows API if GUI object not available
                            try {
                                RefreshWindowDirectly(guiHwnd)
                            } catch {
                                ; Direct refresh also failed
                            }
                        }
                        Sleep(25) ; Reduced processing time between windows
                    } else {
                        ; Window no longer exists, remove from activeGuis
                        try {
                            if (activeGuis.Has(guiHwnd)) {
                                activeGuis.Delete(guiHwnd)
                            }
                        } catch {
                            ; Failed to remove from activeGuis
                        }
                    }
                } catch Error as e {
                    ; Individual window refresh failed, continue with next
                }
            }
        }
        
    } catch Error as e {
        ; Complete refresh operation failed
    } finally {
        isRefreshing := false
    }
}

/**
 * Enhanced GUI refresh with better error handling and validation
 * @param {Object} gui - GUI object to refresh
 * @param {String} windowType - Type description for debugging
 * @return {Void}
 */
RefreshGUITheme(gui, windowType := "") {
    ; Enhanced validation with multiple checks
    if (!IsObject(gui)) {
        return
    }
    
    ; Check if GUI has Hwnd property and it's valid
    try {
        if (!gui.HasOwnProp("Hwnd") || !gui.Hwnd) {
            return
        }
    } catch {
        return
    }
    
    ; Verify window still exists and is valid
    try {
        if (!WinExist("ahk_id " gui.Hwnd) || !DllCall("IsWindow", "Ptr", gui.Hwnd)) {
            return
        }
    } catch {
        return
    }
    
    try {
        ; Apply dark mode to the GUI window itself with validation
        try {
            ApplyDarkModeToGUI(gui, windowType)
        } catch Error as e {
            ; ApplyDarkModeToGUI failed, continue with other operations
        }
        
        ; Get current theme colors
        colors := GetThemeColors()
        isDark := IsWindowsDarkMode()
        
        ; Refresh all controls in the GUI with enhanced error handling
        try {
            RefreshGUIControls(gui, colors, isDark)
        } catch Error as e {
            ; Control refresh failed, continue
        }
        
        ; Apply enhanced text control theming with validation
        try {
            ForceTextControlsRetheming(gui)
        } catch Error as e {
            ; Text control theming failed, continue
        }
        
        ; Final window refresh with validation
        try {
            ; Verify GUI is still valid before redrawing
            if (IsObject(gui) && gui.HasOwnProp("Hwnd") && gui.Hwnd && WinExist("ahk_id " gui.Hwnd)) {
                gui.Redraw()
            }
        } catch Error as e {
            ; Redraw failed, try alternative refresh method
            try {
                if (gui.Hwnd && DllCall("IsWindow", "Ptr", gui.Hwnd)) {
                    DllCall("InvalidateRect", "Ptr", gui.Hwnd, "Ptr", 0, "Int", 1)
                    DllCall("UpdateWindow", "Ptr", gui.Hwnd)
                }
            } catch {
                ; Alternative refresh also failed, skip
            }
        }
        
        ; Force complete repaint with enhanced validation
        try {
            if (gui.Hwnd && DllCall("IsWindow", "Ptr", gui.Hwnd)) {
                DllCall("RedrawWindow", "Ptr", gui.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x0001 | 0x0004 | 0x0010) ; RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN
            }
        } catch Error as e {
            ; RedrawWindow failed, use simpler refresh
            try {
                if (gui.Hwnd && DllCall("IsWindow", "Ptr", gui.Hwnd)) {
                    DllCall("InvalidateRect", "Ptr", gui.Hwnd, "Ptr", 0, "Int", 1)
                }
            } catch {
                ; All refresh methods failed, continue silently
            }
        }
        
    } catch Error as e {
        ; Complete GUI refresh failed - continue silently to prevent crashes
    }
}

/**
 * Enhanced control refresh with better enumeration
 */
RefreshGUIControls(gui, colors, isDark) {
    try {
        ; Use a more reliable method to refresh controls
        ; Send refresh messages to all child windows
        try {
            DllCall("EnumChildWindows", "Ptr", gui.Hwnd, "Ptr", CallbackCreate(RefreshChildControl.Bind(colors, isDark), "F", 2), "Ptr", 0)
        } catch {
            ; Enumeration failed - try refreshing known global controls
            RefreshKnownGlobalControls(colors, isDark)
        }
        
    } catch Error as e {
        ; Control refresh failed
    }
}

/**
 * Enhanced child control refresh callback
 */
RefreshChildControl(colors, isDark, hwnd, lParam) {
    try {
        ; Get control class name
        className := DllCall("GetClassName", "Ptr", hwnd, "Str", "", "Int", 256, "Str")
        
        ; Create a temporary control object to apply theming
        ; This is a workaround since we don't have direct access to the AutoHotkey control object
        try {
            ; Apply basic theming using Windows API
            if (isDark) {
                switch className {
                    case "Edit":
                        DllCall("uxtheme\SetWindowTheme", "Ptr", hwnd, "WStr", "DarkMode_CFD", "Ptr", 0)
                    case "Button":
                        DllCall("uxtheme\SetWindowTheme", "Ptr", hwnd, "WStr", "DarkMode_Explorer", "Ptr", 0)
                    case "ComboBox":
                        DllCall("uxtheme\SetWindowTheme", "Ptr", hwnd, "WStr", "DarkMode_CFD", "Ptr", 0)
                    case "SysListView32":
                        DllCall("uxtheme\SetWindowTheme", "Ptr", hwnd, "WStr", "DarkMode_Explorer", "Ptr", 0)
                }
            } else {
                DllCall("uxtheme\SetWindowTheme", "Ptr", hwnd, "Ptr", 0, "Ptr", 0)
            }
            
            ; Force control redraw
            DllCall("InvalidateRect", "Ptr", hwnd, "Ptr", 0, "Int", 1)
            DllCall("UpdateWindow", "Ptr", hwnd)
            
        } catch {
            ; Individual control theming failed
        }
        
    } catch {
        ; Control processing failed
    }
    
    return true ; Continue enumeration
}

/**
 * Enhanced fallback method to refresh known global controls with better error handling
 * @param {Map} colors - Theme color map
 * @param {Boolean} isDark - Dark mode flag
 */
RefreshKnownGlobalControls(colors, isDark) {
    ; Try to refresh common control variables if they exist in global scope
    global ShortcutsText, CategoryFilter, AddAppButton, PreferencesButton, CloseButton, HelpButton, ExportButton, FilterLabel, ShortcutsHeader

    ; Create an array of the actual control objects and their types
    controlList := [
        {obj: ShortcutsText, type: "Edit"},
        {obj: CategoryFilter, type: "DropDownList"},
        {obj: AddAppButton, type: "Button"},
        {obj: PreferencesButton, type: "Button"},
        {obj: CloseButton, type: "Button"},
        {obj: HelpButton, type: "Button"},
        {obj: ExportButton, type: "Button"},
        {obj: FilterLabel, type: "Text"},
        {obj: ShortcutsHeader, type: "Text"}
    ]

    ; Loop through the list and apply the theme safely
    for controlRef in controlList {
        try {
            ; Enhanced validation before processing each control
            if (IsObject(controlRef.obj) && controlRef.obj.HasOwnProp("Hwnd") && controlRef.obj.Hwnd) {
                ; Verify the control's window still exists
                if (DllCall("IsWindow", "Ptr", controlRef.obj.Hwnd)) {
                    ApplyControlTheming(controlRef.obj, controlRef.type, colors, isDark)
                    Sleep(10)  ; Small delay to prevent overwhelming the system
                }
            }
        } catch {
            ; Control might not exist or be accessible - continue with next
        }
    }
}


/**
 * Enhanced theme monitoring setup with better stability
 */
SetupEnhancedThemeMonitoring() {
    ; Start with normal frequency monitoring
    SetTimer(CheckThemeChange, 2000)  ; Check every 2 seconds for better stability
    
    ; Listen for Windows theme change messages with better handling
    OnMessage(0x001A, HandleThemeChangeMessage) ; WM_WININICHANGE
    OnMessage(0x031A, HandleThemeChangeMessage) ; WM_THEMECHANGED
    OnMessage(0x02B1, HandleThemeChangeMessage) ; WM_WTSSESSION_CHANGE
}

/**
 * Enhanced message handler for theme changes
 */
HandleThemeChangeMessage(wParam, lParam, msg, hwnd) {
    ; Add delay to allow Windows to complete theme change
    SetTimer(CheckThemeChange, -500)

    ; Also schedule a delayed full refresh
    ScheduleFullThemeRefresh()
}

/**
 * Schedules a delayed, full refresh of the entire application theme.
 * This is the primary function to call after a theme change is detected.
 */
ScheduleFullThemeRefresh() {
    SetTimer(ScheduleThemeRefreshCallback, -1000)
}

ScheduleThemeRefreshCallback() {
    ; ForceThemeCheck is the main entry point and will handle
    ; calling the other refresh functions if a change is detected.
    ForceThemeCheck()
}

/**
 * Enhanced theme monitoring with more frequent checks during dialog operations
 * @return {Void}
 */
StartIntensiveThemeMonitoring() {
    global intensiveThemeMonitoring
    ; Increase monitoring frequency when dialogs are active
    intensiveThemeMonitoring := true
    SetTimer(CheckThemeChange, 500)  ; Reduced from 250ms for better stability
}

/**
 * Return to normal theme monitoring frequency
 * @return {Void}
 */
RestoreNormalThemeMonitoring() {
    global intensiveThemeMonitoring
    ; Return to normal 2-second monitoring
    intensiveThemeMonitoring := false
    SetTimer(CheckThemeChange, 2000)
}

/**
 * Helper function to get GUI object from window handle
 * @param {Integer} hwnd - Window handle
 * @return {Object|Boolean} GUI object or false if not found
 */
GuiFromHwnd(hwnd) {
    ; This is a workaround since AutoHotkey v2 doesn't have a direct way
    ; to get GUI object from HWND. We'll check common GUI references.
    global MyGui
    
    try {
        if (MyGui && MyGui.Hwnd = hwnd) {
            return MyGui
        }
    } catch {
        ; MyGui might not be initialized
    }
    
    ; For other GUIs, we can't easily retrieve the object from HWND
    ; but we can still apply basic theming using Windows API
    return false
}

/**
 * Direct window refresh using Windows API when GUI object is not available
 */
RefreshWindowDirectly(hwnd) {
    try {
        isDark := IsWindowsDarkMode()
        colors := GetThemeColors()
        
        ; Apply window-level theming
        if (isDark) {
            try {
                DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", 20, "Int*", 1, "UInt", 4)
                DllCall("uxtheme\SetWindowTheme", "Ptr", hwnd, "WStr", "DarkMode_Explorer", "Ptr", 0)
            } catch {
                ; API calls failed
            }
        } else {
            try {
                DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "UInt", 20, "Int*", 0, "UInt", 4)
                DllCall("uxtheme\SetWindowTheme", "Ptr", hwnd, "Ptr", 0, "Ptr", 0)
            } catch {
                ; API calls failed
            }
        }
        
        ; Force window repaint
        DllCall("InvalidateRect", "Ptr", hwnd, "Ptr", 0, "Int", 1)
        DllCall("UpdateWindow", "Ptr", hwnd)
        
        ; Refresh all child controls
        DllCall("EnumChildWindows", "Ptr", hwnd, "Ptr", CallbackCreate(RefreshChildControlDirect.Bind(isDark), "F", 2), "Ptr", 0)
        
    } catch {
        ; Direct refresh failed
    }
}

/**
 * Direct child control refresh callback
 */
RefreshChildControlDirect(isDark, childHwnd, lParam) {
    try {
        className := DllCall("GetClassName", "Ptr", childHwnd, "Str", "", "Int", 256, "Str")
        
        if (isDark) {
            switch className {
                case "Edit":
                    DllCall("uxtheme\SetWindowTheme", "Ptr", childHwnd, "WStr", "DarkMode_CFD", "Ptr", 0)
                case "Button":
                    DllCall("uxtheme\SetWindowTheme", "Ptr", childHwnd, "WStr", "DarkMode_Explorer", "Ptr", 0)
                case "ComboBox":
                    DllCall("uxtheme\SetWindowTheme", "Ptr", childHwnd, "WStr", "DarkMode_CFD", "Ptr", 0)
                case "SysListView32":
                    DllCall("uxtheme\SetWindowTheme", "Ptr", childHwnd, "WStr", "DarkMode_Explorer", "Ptr", 0)
            }
        } else {
            DllCall("uxtheme\SetWindowTheme", "Ptr", childHwnd, "Ptr", 0, "Ptr", 0)
        }
        
        ; Force control redraw
        DllCall("InvalidateRect", "Ptr", childHwnd, "Ptr", 0, "Int", 1)
        DllCall("UpdateWindow", "Ptr", childHwnd)
        
    } catch {
        ; Control refresh failed
    }
    
    return true
}

;==============================================================================
; FOCUS MANAGEMENT SYSTEM
;==============================================================================

/**
 * Suspend or resume global hotkeys during dialog operations
 * @param {Boolean} suspend - True to suspend hotkeys, false to resume
 * @return {Void}
 */
SuspendGlobalHotkeys(suspend := true) {
    global globalHotkeysSuspended
    
    if (suspend && !globalHotkeysSuspended) {
        try {
            Hotkey("`` & 2", "Off")
            Hotkey("`` & 3", "Off") 
            globalHotkeysSuspended := true
        } catch {
            ; Hotkeys might not be registered yet
        }
    } else if (!suspend && globalHotkeysSuspended) {
        try {
            Hotkey("`` & 2", "On")
            Hotkey("`` & 3", "On")
            globalHotkeysSuspended := false
        } catch {
            ; Hotkeys might have been removed
        }
    }
}

/**
 * Start monitoring dialog focus to prevent focus theft by other windows
 * @param {Object} gui - GUI object to monitor for focus
 * @param {Integer} intervalMs - Check interval in milliseconds (default 500)
 * @return {Void}
 */
StartFocusMonitoring(gui, intervalMs := 500) {
    global activeFocusTimers
    
    guiHwnd := gui.Hwnd
    activeFocusTimers[guiHwnd] := SetTimer(CheckAndRestoreFocus.Bind(guiHwnd), intervalMs)
}

/**
 * Check if monitored dialog has focus and restore if necessary
 * @param {Integer} hwnd - Window handle to check and restore focus for
 * @return {Void}
 */
CheckAndRestoreFocus(hwnd) {
    global activeGuis
    
    if (WinExist("ahk_id " hwnd)) {
        if (!WinActive("ahk_id " hwnd)) {
            ; Check if another dialog should have focus instead
            hasOtherDialog := false
            for otherHwnd, callback in activeGuis {
                if (otherHwnd != hwnd && WinActive("ahk_id " otherHwnd)) {
                    hasOtherDialog := true
                    break
                }
            }
            
            ; Restore focus only if no other dialog is active
            if (!hasOtherDialog) {
                try {
                    WinActivate("ahk_id " hwnd)
                } catch {
                    ; Window might be in the process of closing
                }
            }
        }
    } else {
        ; Window no longer exists, stop monitoring
        StopFocusMonitoring(hwnd)
    }
}

/**
 * Stop focus monitoring for a specific dialog window
 * @param {Integer} guiHwnd - Window handle to stop monitoring
 * @return {Void}
 */
StopFocusMonitoring(guiHwnd) {
    global activeFocusTimers
    
    if (activeFocusTimers.Has(guiHwnd)) {
        SetTimer(activeFocusTimers[guiHwnd], 0)
        activeFocusTimers.Delete(guiHwnd)
    }
}

/**
 * Enhanced CreateManagedDialog with theme stability
 * @param {String} options - GUI creation options string
 * @param {String} title - Window title text
 * @param {Boolean} enableFocusMonitoring - Enable automatic focus monitoring
 * @return {Object} Enhanced GUI object with focus management features
 */
CreateManagedDialog(options, title, enableFocusMonitoring := true) {
    ; Pre-check theme before creating dialog
    ForceThemeCheck()
    
    ; Start intensive theme monitoring during dialog creation
    StartIntensiveThemeMonitoring()
    
    ; Suspend global hotkeys to prevent interference
    SuspendGlobalHotkeys(true)
    
    try {
        ; Create GUI with retries for stability
        gui := false
        attempts := 0
        
        while (!gui && attempts < 3) {
            try {
                gui := Gui(options, title)
                break
            } catch {
                attempts++
                Sleep(100)
            }
        }
        
        if (!gui) {
            throw Error("Failed to create GUI after multiple attempts")
        }
        
        ; Apply theming with error handling
        try {
            ApplyDarkModeToGUI(gui, title)
            
            ; Force immediate theme application
            colors := GetThemeColors()
            gui.BackColor := colors["background"]
            
            ; Set font with proper color
            textColor := Format("c0x{:06X}", colors["textColor"])
            gui.SetFont("s10 " textColor)
            
            ; Apply enhanced text control theming
            ForceTextControlsRetheming(gui)
            
        } catch Error as e {
            ; Continue with basic GUI if theming fails
        }
        
        ; Set up close event handler
        gui.OnEvent("Close", (*) => HandleDialogCloseWithThemeCleanup(gui.Hwnd))
        
        ; Enable focus monitoring if requested
        if (enableFocusMonitoring) {
            SetTimer(StartFocusMonitoringDelayed.Bind(gui), -200)
        }
        
        return gui
        
    } catch Error as e {
        SuspendGlobalHotkeys(false)
        RestoreNormalThemeMonitoring()
        throw e
    }
}

/**
 * Start focus monitoring with delay to allow dialog to fully initialize
 * @param {Object} gui - GUI object to monitor
 * @return {Void}
 */
StartFocusMonitoringDelayed(gui, *) {
    if (WinExist("ahk_id " gui.Hwnd)) {
        StartFocusMonitoring(gui)
    }
}

/**
 * Enhanced dialog close handler with theme monitoring cleanup
 * @param {Integer} guiHwnd - Window handle being closed
 * @return {Void}
 */
HandleDialogCloseWithThemeCleanup(guiHwnd, *) {
    global activeGuis
    
    ; Stop focus monitoring for this dialog
    StopFocusMonitoring(guiHwnd)
    
    ; If this was the last dialog, restore normal theme monitoring
    if (activeGuis.Count <= 1) {
        SuspendGlobalHotkeys(false)
        RestoreNormalThemeMonitoring()
    }
}

/**
 * Force window to foreground using advanced Windows API methods
 * @param {Integer} hwnd - Window handle to bring to foreground
 * @return {Void}
 */
ForceWindowToForeground(hwnd) {
    foregroundHwnd := DllCall("GetForegroundWindow", "Ptr")
    
    if (foregroundHwnd != hwnd) {
        ; Get thread IDs for thread input attachment
        foregroundThread := DllCall("GetWindowThreadProcessId", "Ptr", foregroundHwnd, "Ptr", 0, "UInt")
        currentThread := DllCall("GetCurrentThreadId", "UInt")
        
        ; Attach thread inputs to bypass Windows focus restrictions
        if (foregroundThread != currentThread) {
            DllCall("AttachThreadInput", "UInt", currentThread, "UInt", foregroundThread, "Int", 1)
        }
        
        ; Set foreground and focus
        DllCall("SetForegroundWindow", "Ptr", hwnd)
        DllCall("SetFocus", "Ptr", hwnd)
        
        ; Detach thread inputs
        if (foregroundThread != currentThread) {
            DllCall("AttachThreadInput", "UInt", currentThread, "UInt", foregroundThread, "Int", 0)
        }
    }
}

/**
 * Register GUI for Escape key handling with enhanced focus management
 * @param {Object} gui - GUI object to register for Escape handling
 * @param {Function} closeCallback - Function to call when Escape is pressed
 * @return {Void}
 */
RegisterGuiForEscapeHandling(gui, closeCallback) {
    global activeGuis
    
    guiHwnd := gui.Hwnd
    activeGuis[guiHwnd] := closeCallback
    
    ; Clean up function for when dialog closes
    SafeRemoveGui(*) {
        StopFocusMonitoring(guiHwnd)
        if activeGuis.Has(guiHwnd)
            activeGuis.Delete(guiHwnd)
        
        ; Resume global hotkeys if no dialogs remain
        if (activeGuis.Count = 0) {
            SuspendGlobalHotkeys(false)
            RestoreNormalThemeMonitoring()
        }
    }
    
    gui.OnEvent("Close", SafeRemoveGui)
    
    ; Set up window-specific Escape hotkey
    HotIfWinActive("ahk_id " guiHwnd)
    Hotkey "Escape", closeCallback
    HotIf()
}

;==============================================================================
; CORE UTILITY FUNCTIONS
;==============================================================================

/**
 * Display debug or status message based on application mode
 * @param {String} msg - Message to display to user
 * @return {Void}
 */
ShowMessage(msg) {
    if (!quietMode && debugMode) {
        MsgBox(msg)
    }
}

/**
 * Convert process executable name to user-friendly application name
 * @param {String} processName - Process executable name (e.g., "chrome.exe")
 * @return {String} Human-readable application name (e.g., "Google Chrome")
 */
GetFriendlyAppName(processName) {
    ; Convert to lowercase for case-insensitive matching
    lowerProcessName := StrLower(processName)
    
    ; Comprehensive application name mappings
    appNames := Map(
        ; Web Browsers
        "chrome.exe", "Google Chrome",
        "firefox.exe", "Mozilla Firefox", 
        "msedge.exe", "Microsoft Edge",
        "opera.exe", "Opera Browser",
        "brave.exe", "Brave Browser",
        "vivaldi.exe", "Vivaldi Browser",
        "iexplore.exe", "Internet Explorer",
        
        ; Text Editors & IDEs
        "notepad.exe", "Notepad",
        "notepad++.exe", "Notepad++",
        "code.exe", "Visual Studio Code",
        "devenv.exe", "Visual Studio",
        "sublime_text.exe", "Sublime Text",
        "atom.exe", "Atom Editor",
        
        ; Microsoft Office Suite
        "winword.exe", "Microsoft Word",
        "excel.exe", "Microsoft Excel", 
        "powerpnt.exe", "Microsoft PowerPoint",
        "msaccess.exe", "Microsoft Access",
        "outlook.exe", "Microsoft Outlook",
        "onenote.exe", "Microsoft OneNote",
        "teams.exe", "Microsoft Teams",
        
        ; Windows System Applications
        "explorer.exe", "Windows Explorer",
        "cmd.exe", "Command Prompt",
        "powershell.exe", "Windows PowerShell",
        "pwsh.exe", "PowerShell Core",
        "taskmgr.exe", "Task Manager",
        "mmc.exe", "Microsoft Management Console",
        "regedit.exe", "Registry Editor",
        "calc.exe", "Calculator",
        "mspaint.exe", "Paint",
        "wordpad.exe", "WordPad",
        
        ; Adobe Creative Suite
        "photoshop.exe", "Adobe Photoshop",
        "illustrator.exe", "Adobe Illustrator", 
        "indesign.exe", "Adobe InDesign",
        "acrobat.exe", "Adobe Acrobat",
        "acrod32.exe", "Adobe Acrobat Reader",
        "premiere.exe", "Adobe Premiere Pro",
        "afterfx.exe", "Adobe After Effects",
        
        ; Communication & Social
        "spotify.exe", "Spotify",
        "discord.exe", "Discord",
        "slack.exe", "Slack",
        "zoom.exe", "Zoom",
        "skype.exe", "Skype",
        "telegram.exe", "Telegram",
        "whatsapp.exe", "WhatsApp",
        
        ; Gaming Platforms
        "steam.exe", "Steam",
        "epic games launcher.exe", "Epic Games Launcher",
        "origin.exe", "EA Origin",
        "uplay.exe", "Ubisoft Connect",
        
        ; Media Players
        "vlc.exe", "VLC Media Player",
        "wmplayer.exe", "Windows Media Player",
        "itunes.exe", "Apple iTunes",
        "foobar2000.exe", "foobar2000",
        
        ; File Management & Utilities
        "7zfm.exe", "7-Zip File Manager",
        "winrar.exe", "WinRAR",
        "filezilla.exe", "FileZilla",
        "putty.exe", "PuTTY",
        "winscp.exe", "WinSCP",
        
        ; Development Tools
        "javaw.exe", "Java Application",
        "java.exe", "Java Application",
        "python.exe", "Python",
        "node.exe", "Node.js"
    )
    
    ; Check if we have a friendly name mapping
    if (appNames.Has(lowerProcessName)) {
        return appNames[lowerProcessName]
    }
    
    ; If no mapping found, clean up the executable name
    cleanName := StrReplace(processName, ".exe", "", false)
    if (cleanName != "") {
        ; Replace underscores/hyphens with spaces and capitalize
        cleanName := StrReplace(cleanName, "_", " ")
        cleanName := StrReplace(cleanName, "-", " ")
        cleanName := StrTitle(cleanName)
        return cleanName
    }
    
    ; Final fallback to original process name
    return processName
}

/**
 * Generate display name combining friendly name with executable
 * @param {String} processName - Process executable name
 * @return {String} Display name with executable in parentheses when applicable
 */
GetAppDisplayName(processName) {
    friendlyName := GetFriendlyAppName(processName)
    
    ; Show both friendly name and executable if they differ significantly
    if (StrLower(friendlyName) != StrLower(processName) && !InStr(friendlyName, processName)) {
        return friendlyName " (" processName ")"
    }
    
    return friendlyName
}

/**
 * Join array elements with specified delimiter
 * @param {Array} arr - Array of strings to join
 * @param {String} delimiter - Separator string between elements
 * @return {String} Joined string result
 */
Join(arr, delimiter) {
    result := ""
    for i, item in arr {
        if (i > 1)
            result .= delimiter
        result .= item
    }
    return result
}

/**
 * Check if a window control exists
 * @param {String} Control - Control identifier to check
 * @param {String} WinTitle - Window title (optional, defaults to active window)
 * @return {Boolean} True if control exists and is accessible
 */
ControlExist(Control, WinTitle := "") {
    try {
        ControlGetPos(, , , , Control, WinTitle)
        return true
    } catch {
        return false
    }
}

/**
 * Compare two semantic version strings
 * @param {String} version1 - First version string (e.g., "2.5.1")
 * @param {String} version2 - Second version string (e.g., "2.4.0")  
 * @return {Integer} 1 if v1>v2, -1 if v1<v2, 0 if equal
 */
CompareVersions(version1, version2) {
    v1Parts := StrSplit(version1, ".")
    v2Parts := StrSplit(version2, ".")
    maxLength := Max(v1Parts.Length, v2Parts.Length)
    
    Loop (maxLength > 0 ? maxLength : 0) {
        v1Part := A_Index <= v1Parts.Length ?
		Integer(v1Parts[A_Index]) : 0
        v2Part := A_Index <= v2Parts.Length ?
		Integer(v2Parts[A_Index]) : 0
        if (v1Part > v2Part)
            return 1
        else if (v1Part < v2Part)
            return -1
    }
    return 0
}

/**
 * Debug utility to display object properties
 * @param {Object} obj - Object to inspect and display
 * @return {String} Formatted property list for debugging
 */
PrintObjectProps(obj) {
    result := ""
    for key, value in obj.OwnProps() {
        result .= key " = " value "`n"
    }
    return result
}

; --- PASTE THE HELPER FUNCTION HERE ---
AddToArrayIfMissing(arr, item) {
    for existingItem in arr {
        if (existingItem = item)
            return
    }
    arr.Push(item)
}

;==============================================================================
; INSTALLATION SYSTEM
;==============================================================================

/**
 * Verify all required files and dependencies are available
 * @return {Boolean} True if all requirements are met for operation
 */
CheckProgramRequirements() {
    global REQUIRED_FILES
    missingFiles := []
    
    ; For compiled versions, check temp directory where FileInstall extracts
    ; For non-compiled versions, check script directory
    if (A_IsCompiled) {
        tempDir := A_Temp "\Keyfinder_Temp"
        for file in REQUIRED_FILES {
            if (!FileExist(tempDir "\" file)) {
                missingFiles.Push(file)
            }
        }
    } else {
        ; Check for required files in script directory (non-compiled)
        for file in REQUIRED_FILES {
            if (!FileExist(A_ScriptDir "\" file)) {
                missingFiles.Push(file)
            }
        }
    }
    
    ; Verify SQLite availability (checks multiple locations)
    if (!SQLiteAvailable()) {
        missingFiles.Push("SQLite library")
    }
    
    ; Handle missing files for compiled versions
    if (missingFiles.Length > 0 && A_IsCompiled) {
        try {
            Sleep(1000) ; Allow FileInstall extraction to complete
            
            ; Re-check after allowing time for extraction
            tempDir := A_Temp "\Keyfinder_Temp"
            stillMissing := []
            for file in REQUIRED_FILES {
                if (!FileExist(tempDir "\" file)) {
                    stillMissing.Push(file)
                }
            }
            
            if (stillMissing.Length > 0) {
                MsgBox("Missing required files: " Join(stillMissing, ", ") 
                     . "`n`nThe program may not function correctly.", "Requirements Check", 48)
                return false
            }
        } catch Error as e {
            MsgBox("Error extracting required files: " e.Message, "Requirements Error", 16)
            return false
        }
    } else if (missingFiles.Length > 0) {
        MsgBox("Missing required files: " Join(missingFiles, ", "), "Requirements Check", 48)
        return false
    }
    
    return true
}

/**
 * Get the default installation path in user's Documents folder
 * @return {String} Full installation directory path
 */
GetInstallationPath() {
    global SCRIPT_NAME
    documentsPath := EnvGet("USERPROFILE") "\Documents"
    return documentsPath "\" SCRIPT_NAME
}

/**
 * Analyze current installation status and version information
 * @return {Object} Installation status: {isFirstRun, needsUpgrade, installedVersion}
 */
CheckInstallationStatus() {
    global INSTALL_DIR, SETTINGS_FILE, SCRIPT_VERSION
    
    ; Check if running from installation directory
    if (A_ScriptDir != INSTALL_DIR) {
        return {isFirstRun: true, needsUpgrade: false, installedVersion: ""}
    }
    
    ; Check if settings file exists
    if (!FileExist(SETTINGS_FILE)) {
        return {isFirstRun: true, needsUpgrade: false, installedVersion: ""}
    }
    
    ; Compare installed version with current version
    try {
        installedVersion := IniRead(SETTINGS_FILE, "General", "Version", "0.0")
        versionComparison := CompareVersions(SCRIPT_VERSION, installedVersion)
        
        if (versionComparison != 0) {
            return {isFirstRun: false, needsUpgrade: true, installedVersion: installedVersion}
        } else {
            return {isFirstRun: false, needsUpgrade: false, installedVersion: installedVersion}
        }
    } catch {
        return {isFirstRun: true, needsUpgrade: false, installedVersion: ""}
    }
}

/**
 * Present version upgrade options to user with detailed information
 * @param {String} currentVersion - Version being installed
 * @param {String} installedVersion - Currently installed version
 * @return {Boolean} True if user approves the upgrade/downgrade
 */
AskForUpgrade(currentVersion, installedVersion) {
    global SCRIPT_NAME
    versionComparison := CompareVersions(currentVersion, installedVersion)
    
    if (versionComparison < 0) {
        ; Downgrade warning
        msg := "Version Downgrade Detected!`n`n"
        msg .= "Currently Installed: v" installedVersion "`n"
        msg .= "Version to Install: v" currentVersion "`n`n"
        msg .= "You are trying to install an older version. Continue?"
        result := MsgBox(msg, "Downgrade Warning", 52)
    } else if (versionComparison = 0) {
        ; Reinstall confirmation
        msg := "Same Version Detected!`n`n"
        msg .= "Version " currentVersion " is already installed.`n`n"
        msg .= "Would you like to reinstall/repair the current version?"
        result := MsgBox(msg, "Reinstall Confirmation", 36)
    } else {
        ; Normal upgrade
        msg := "A newer version of " SCRIPT_NAME " is being installed!`n`n"
        msg .= "Currently Installed: v" installedVersion "`n"
        msg .= "New Version: v" currentVersion "`n`n"
        msg .= "Your shortcuts database and settings will be preserved."
        result := MsgBox(msg, "Upgrade Available", 36)
    }
    
    return (result = "Yes")
}

/**
 * Execute comprehensive upgrade/downgrade/reinstall process
 * @param {String} installedVersion - Currently installed version for comparison
 * @return {Boolean} True if upgrade process completed successfully
 */
PerformUpgrade(installedVersion) {
    global INSTALL_DIR, SCRIPT_NAME, SETTINGS_FILE, SCRIPT_VERSION
    
    versionComparison := CompareVersions(SCRIPT_VERSION, installedVersion)
    operationType := versionComparison > 0 ? "Upgrading" : 
                     versionComparison < 0 ? "Downgrading" : "Reinstalling"
    
    ; Create progress dialog for user feedback
; Use the existing function to create the dialog
progressDialog := CreateProgressDialog()

; Get references to all the created components
ProgressGui := progressDialog.gui
ProgressBar := progressDialog.bar
StatusText := progressDialog.status
TitleText := progressDialog.title ; Get the new reference to the title text

; Customize the dialog for the upgrade process
ProgressGui.Title := operationType " " SCRIPT_NAME
TitleText.Text := operationType " " SCRIPT_NAME " to v" SCRIPT_VERSION "..." ; Use the direct reference
StatusText.Text := "Starting " operationType "..."
    try {
        ; Step 1: Backup user data with progress indication
        if (ProgressGui) {
            ProgressBar.Value := 20
            StatusText.Text := "Backing up user data..."
            Sleep(500)
        }
        
        tempBackupDir := A_Temp "\" SCRIPT_NAME "_upgrade_backup"
        if (DirExist(tempBackupDir)) {
            DirDelete(tempBackupDir, 1)
        }
        DirCreate(tempBackupDir)
        
        ; Backup critical user files
        if (FileExist(SETTINGS_FILE)) {
            FileCopy(SETTINGS_FILE, tempBackupDir "\settings.ini", 1)
        }
        if (FileExist(INSTALL_DIR "\appinfo.db")) {
            FileCopy(INSTALL_DIR "\appinfo.db", tempBackupDir "\appinfo.db", 1)
        }
        if (DirExist(INSTALL_DIR "\Backups")) {
            DirCopy(INSTALL_DIR "\Backups", tempBackupDir "\Backups", 1)
        }
        if (DirExist(INSTALL_DIR "\Exports")) {
            DirCopy(INSTALL_DIR "\Exports", tempBackupDir "\Exports", 1)
        }
        
        ; Step 2: Remove old version files
        if (ProgressGui) {
            ProgressBar.Value := 40
            StatusText.Text := "Removing old version files..."
            Sleep(500)
        }
        
        oldFiles := ["sqlite3.dll"]
        for file in oldFiles {
            if (FileExist(INSTALL_DIR "\" file)) {
                try FileDelete(INSTALL_DIR "\" file)
            }
        }
        
        ; Step 3: Install new version files
        if (ProgressGui) {
            ProgressBar.Value := 60
            StatusText.Text := "Installing new version..."
            Sleep(500)
        }
        
        sourceExe := A_ScriptFullPath
        targetExe := INSTALL_DIR "\" A_ScriptName
        
        if (sourceExe != targetExe) {
            FileCopy(sourceExe, targetExe, 1)
        }
        
        requiredFiles := ["sqlite3.dll"]
        for file in requiredFiles {
            sourcePath := A_ScriptDir "\" file
            targetPath := INSTALL_DIR "\" file
            
            if (FileExist(sourcePath) && sourcePath != targetPath) {
                FileCopy(sourcePath, targetPath, 1)
            }
        }
        
        ; Handle database file carefully (preserve existing data)
        if (!FileExist(INSTALL_DIR "\appinfo.db") || versionComparison = 0) {
            sourcePath := A_ScriptDir "\appinfo.db"
            targetPath := INSTALL_DIR "\appinfo.db"
            
            if (FileExist(sourcePath) && sourcePath != targetPath) {
                FileCopy(sourcePath, targetPath, 0)
            }
        }
        
        ; Step 4: Restore user data
        if (ProgressGui) {
            ProgressBar.Value := 80
            StatusText.Text := "Restoring user data..."
            Sleep(500)
        }
        
        if (FileExist(tempBackupDir "\settings.ini")) {
            FileCopy(tempBackupDir "\settings.ini", SETTINGS_FILE, 1)
        }
        if (FileExist(tempBackupDir "\appinfo.db")) {
            FileCopy(tempBackupDir "\appinfo.db", INSTALL_DIR "\appinfo.db", 1)
        }
        if (DirExist(tempBackupDir "\Backups")) {
            DirCopy(tempBackupDir "\Backups", INSTALL_DIR "\Backups", 1)
        }
        if (DirExist(tempBackupDir "\Exports")) {
            DirCopy(tempBackupDir "\Exports", INSTALL_DIR "\Exports", 1)
        }
        
        ; Step 5: Update configuration files
        if (ProgressGui) {
            ProgressBar.Value := 90
            StatusText.Text := "Updating configuration..."
            Sleep(500)
        }
        
        IniWrite(SCRIPT_VERSION, SETTINGS_FILE, "General", "Version")
        IniWrite(FormatTime(, "yyyy-MM-dd HH:mm:ss"), SETTINGS_FILE, "General", operationType "Date")
        IniWrite(installedVersion, SETTINGS_FILE, "General", "PreviousVersion")
        
        ; Step 6: Cleanup and completion
        if (ProgressGui) {
            ProgressBar.Value := 100
            StatusText.Text := operationType " complete!"
            Sleep(1000)
            ProgressGui.Destroy()
        }
        
        if (DirExist(tempBackupDir)) {
            DirDelete(tempBackupDir, 1)
        }
        
        ; Display appropriate completion message
        if (versionComparison > 0) {
            msg := SCRIPT_NAME " has been successfully upgraded!`n`n"
            msg .= "Previous Version: v" installedVersion "`n"
            msg .= "New Version: v" SCRIPT_VERSION "`n`n"
            msg .= "Your shortcuts database and settings have been preserved."
        } else if (versionComparison < 0) {
            msg := SCRIPT_NAME " has been downgraded.`n`n"
            msg .= "Previous Version: v" installedVersion "`n"
            msg .= "Current Version: v" SCRIPT_VERSION "`n`n"
            msg .= "Note: Some features may no longer be available."
        } else {
            msg := SCRIPT_NAME " has been successfully reinstalled!`n`n"
            msg .= "Version: v" SCRIPT_VERSION "`n`n"
            msg .= "Any corrupted files have been repaired."
        }
        
        MsgBox(msg, operationType " Complete", 64)
        return true
        
    } catch Error as e {
        if (ProgressGui) {
            ProgressGui.Destroy()
        }
        
        ; Attempt to restore from backup on failure
        try {
            if (FileExist(tempBackupDir "\settings.ini")) {
                FileCopy(tempBackupDir "\settings.ini", SETTINGS_FILE, 1)
            }
            if (FileExist(tempBackupDir "\appinfo.db")) {
                FileCopy(tempBackupDir "\appinfo.db", INSTALL_DIR "\appinfo.db", 1)
            }
        } catch {
            ; Backup restoration failed
        }
        
        MsgBox(operationType " failed: " e.Message, operationType " Error", 16)
        return false
    }
}

/**
 * Create complete installation directory structure
 * @return {Boolean} True if directory structure created successfully
 */
CreateInstallationDirectory() {
    global INSTALL_DIR
    
    try {
        if (!DirExist(INSTALL_DIR)) {
            DirCreate(INSTALL_DIR)
        }
        
        ; Create required subdirectories
        subdirs := ["Backups", "Exports", "Logs"]
        for subdir in subdirs {
            subdirPath := INSTALL_DIR "\" subdir
            if (!DirExist(subdirPath)) {
                DirCreate(subdirPath)
            }
        }
        
        return true
    } catch Error as e {
        MsgBox("Error creating installation directory: " e.Message, "Installation Error", 16)
        return false
    }
}

/**
 * Copy all required program files to installation directory
 * @return {Boolean} True if all files copied successfully
 */
CopyProgramFiles() {
    global INSTALL_DIR, REQUIRED_FILES
    
    try {
        ; Copy main executable
        sourceExe := A_ScriptFullPath
        targetExe := INSTALL_DIR "\" A_ScriptName
        
        if (sourceExe != targetExe) {
            FileCopy(sourceExe, targetExe, 1)
        }
        
        ; Copy required dependency files
        for file in REQUIRED_FILES {
            sourcePath := A_ScriptDir "\" file
            targetPath := INSTALL_DIR "\" file
            
            if (FileExist(sourcePath) && sourcePath != targetPath) {
                FileCopy(sourcePath, targetPath, 1)
            }
        }
        
        return true
    } catch Error as e {
        MsgBox("Error copying program files: " e.Message, "Installation Error", 16)
        return false
    }
}

/**
 * Search for and restore previous user data from backup location
 * @return {Boolean} True if user data was found and restored
 */
RestoreUserDataIfExists() {
    global INSTALL_DIR, SCRIPT_NAME, SETTINGS_FILE, DB_FILE
    
    ; Check for backup directory from previous installations
    backupDir := EnvGet("USERPROFILE") "\Documents\" SCRIPT_NAME "_UserData"
    
    if (!DirExist(backupDir)) {
        return false
    }
    
    restoreMsg := "Previous user data found! Would you like to restore your shortcuts database and settings?"
    result := MsgBox(restoreMsg, "Restore User Data", 36)
    
    if (result != "Yes") {
        return false
    }
    
    try {
        ; Restore configuration files
        if (FileExist(backupDir "\settings.ini")) {
            FileCopy(backupDir "\settings.ini", SETTINGS_FILE, 1)
        }
        
        ; Restore database
        if (FileExist(backupDir "\appinfo.db")) {
            FileCopy(backupDir "\appinfo.db", DB_FILE, 1)
        }
        
        ; Restore backup files
        if (DirExist(backupDir "\Backups")) {
            DirCopy(backupDir "\Backups", INSTALL_DIR "\Backups", 1)
        }
        
        ; Restore export files
        if (DirExist(backupDir "\Exports")) {
            DirCopy(backupDir "\Exports", INSTALL_DIR "\Exports", 1)
        }
        
        ; Clean up temporary backup directory
        DirDelete(backupDir, 1)
        MsgBox("Your data has been successfully restored!", "Data Restored", 64)
        return true
        
    } catch Error as e {
        MsgBox("Error restoring user data: " e.Message, "Restore Error", 48)
        return false
    }
}

/**
 * Create initial settings configuration file with default values
 * @return {Boolean} True if settings file created successfully
 */
CreateSettingsFile() {
    global SETTINGS_FILE, SCRIPT_VERSION
    
    try {
        IniWrite(SCRIPT_VERSION, SETTINGS_FILE, "General", "Version")
        IniWrite(FormatTime(, "yyyy-MM-dd HH:mm:ss"), SETTINGS_FILE, "General", "InstallDate")
        IniWrite("false", SETTINGS_FILE, "General", "StartupEnabled")
        IniWrite("false", SETTINGS_FILE, "General", "DesktopShortcut")
        IniWrite("false", SETTINGS_FILE, "General", "QuietMode")
        IniWrite("false", SETTINGS_FILE, "General", "DebugMode")
        
        return true
    } catch Error as e {
        MsgBox("Error creating settings file: " e.Message, "Installation Error", 16)
        return false
    }
}

/**
 * Prompt user for Windows startup integration preference
 * @return {Boolean} True if user wants startup integration enabled
 */
AskForStartupIntegration() {
    global SCRIPT_NAME
    
    msg := SCRIPT_NAME " has been successfully installed!`n`n"
    msg .= "Would you like " SCRIPT_NAME " to start automatically when Windows starts?"
    result := MsgBox(msg, "Startup Integration", 36)
    
    return (result = "Yes")
}

/**
 * Prompt user for desktop shortcut creation preference
 * @return {Boolean} True if user wants a desktop shortcut created
 */
AskForDesktopShortcut() {
    global SCRIPT_NAME
    
    msg := "Would you like to create a desktop shortcut for " SCRIPT_NAME "?"
    result := MsgBox(msg, "Desktop Shortcut", 36)
    
    return (result = "Yes")
}

/**
 * Configure Windows startup integration through registry
 * @param {Boolean} enable - True to enable startup, false to disable
 * @return {Boolean} True if registry operation completed successfully
 */
SetStartupIntegration(enable) {
    global INSTALL_DIR, SCRIPT_NAME, SETTINGS_FILE
    
    try {
        regKey := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
        executablePath := INSTALL_DIR "\" A_ScriptName
        
        if (enable) {
            RegWrite(executablePath, "REG_SZ", regKey, SCRIPT_NAME)
            IniWrite("true", SETTINGS_FILE, "General", "StartupEnabled")
            ShowMessage("Successfully added to Windows startup.")
        } else {
            try {
                RegDelete(regKey, SCRIPT_NAME)
            } catch {
                ; Registry key might not exist
            }
            IniWrite("false", SETTINGS_FILE, "General", "StartupEnabled")
            ShowMessage("Successfully removed from Windows startup.")
        }
        
        return true
    } catch Error as e {
        MsgBox("Error setting startup integration: " e.Message, "Startup Error", 16)
        return false
    }
}

/**
 * Create or remove desktop shortcut file
 * @param {Boolean} create - True to create shortcut, false to remove
 * @return {Boolean} True if shortcut operation completed successfully
 */
SetDesktopShortcut(create) {
    global INSTALL_DIR, SCRIPT_NAME, SETTINGS_FILE
    
    try {
        desktopPath := EnvGet("USERPROFILE") "\Desktop"
        shortcutPath := desktopPath "\" SCRIPT_NAME ".lnk"
        executablePath := INSTALL_DIR "\" A_ScriptName
        
        if (create) {
            ; Create Windows shortcut using COM object
            shortcut := ComObject("WScript.Shell").CreateShortcut(shortcutPath)
            shortcut.TargetPath := executablePath
            shortcut.WorkingDirectory := INSTALL_DIR
            shortcut.Description := SCRIPT_NAME " - Application Info Utility"
            shortcut.Save()
            
            IniWrite("true", SETTINGS_FILE, "General", "DesktopShortcut")
            ShowMessage("Desktop shortcut created successfully.")
        } else {
            if (FileExist(shortcutPath)) {
                FileDelete(shortcutPath)
            }
            
            IniWrite("false", SETTINGS_FILE, "General", "DesktopShortcut")
            ShowMessage("Desktop shortcut removed successfully.")
        }
        
        return true
    } catch Error as e {
        MsgBox("Error managing desktop shortcut: " e.Message, "Shortcut Error", 16)
        return false
    }
}

/**
 * Create installation progress dialog interface
 * @return {Object} Progress dialog components: {gui, bar, status}
 */
CreateProgressDialog() {
    global SCRIPT_NAME
    
    ProgressGui := Gui("+ToolWindow -MaximizeBox -MinimizeBox", SCRIPT_NAME " Setup")
    ProgressGui.SetFont("s10")
    ApplyDarkModeToGUI(ProgressGui, "Installation Progress")
    colors := GetThemeColors()
    isDark := IsWindowsDarkMode()
    
    titleText := ProgressGui.Add("Text", "x20 y20 w360 center", "Setting up " SCRIPT_NAME "...")
    ApplyControlTheming(titleText, "Text", colors, isDark)
    
    ProgressBar := ProgressGui.Add("Progress", "x20 y50 w360 h20", 0)
    StatusText := ProgressGui.Add("Text", "x20 y80 w360 center", "Initializing...")
    ApplyControlTheming(StatusText, "Text", colors, isDark)
    
    ProgressGui.Show("w400 h120")
    
    ; Add escape handling to progress dialog
    RegisterGuiForEscapeHandling(ProgressGui, (*) => ProgressGui.Destroy())
    
	return {gui: ProgressGui, bar: ProgressBar, status: StatusText, title: titleText}
}

/**
 * Update installation progress display with current status
 * @param {Object} progressDialog - Progress dialog components from CreateProgressDialog
 * @param {Integer} percent - Progress percentage (0-100)
 * @param {String} status - Current operation status message
 * @return {Void}
 */
UpdateProgress(progressDialog, percent, status) {
    if (progressDialog && progressDialog.bar && progressDialog.status) {
        progressDialog.bar.Value := percent
        progressDialog.status.Text := status
        Sleep(100)
    }
}

/**
 * Execute complete first-time installation process
 * @return {Boolean} True if installation completed successfully
 */
PerformInstallation() {
    global INSTALL_DIR, SETTINGS_FILE, SCRIPT_NAME
    
    progress := CreateProgressDialog()
    
    try {
        ; Phase 1: Requirements check
        UpdateProgress(progress, 10, "Checking requirements...")
        if (!CheckProgramRequirements()) {
            progress.gui.Destroy()
            return false
        }
        
        ; Phase 2: Directory creation
        UpdateProgress(progress, 30, "Creating installation directory...")
        if (!CreateInstallationDirectory()) {
            progress.gui.Destroy()
            return false
        }
        
        ; Phase 3: File copying
        UpdateProgress(progress, 50, "Copying program files...")
        if (!CopyProgramFiles()) {
            progress.gui.Destroy()
            return false
        }
        
        ; Phase 4: Data restoration check
        UpdateProgress(progress, 60, "Checking for previous user data...")
        dataRestored := RestoreUserDataIfExists()
        
        ; Phase 5: Configuration setup
        if (!dataRestored) {
            UpdateProgress(progress, 70, "Creating configuration...")
            if (!CreateSettingsFile()) {
                progress.gui.Destroy()
                return false
            }
        } else {
            UpdateProgress(progress, 70, "Configuration restored...")
            Sleep(500)
        }
        
        ; Phase 6: Finalization
        UpdateProgress(progress, 90, "Finalizing installation...")
        Sleep(500)
        
        UpdateProgress(progress, 100, "Installation complete!")
        Sleep(1000)
        
        progress.gui.Destroy()
        
        ; Configure optional components for new installations
        if (!dataRestored && AskForStartupIntegration()) {
            SetStartupIntegration(true)
        }
        
        if (!dataRestored && AskForDesktopShortcut()) {
            SetDesktopShortcut(true)
        }
        
        ; Display completion message
        completionMsg := SCRIPT_NAME " has been successfully "
        completionMsg .= dataRestored ? "reinstalled with your previous data restored" : "installed"
        completionMsg .= " to:`n" INSTALL_DIR
        
        if (!dataRestored) {
            completionMsg .= "`n`nTIP: You can change preferences anytime through the main application."
        }
        
        completionMsg .= "`n`nThe application will now restart from the new location."
        
        MsgBox(completionMsg, "Installation Complete", 64)
        
        ; Restart from installation directory if needed
        if (A_ScriptDir != INSTALL_DIR) {
            newExecutable := INSTALL_DIR "\" A_ScriptName
            if (FileExist(newExecutable)) {
                Run(newExecutable)
                ExitApp()
            }
        }
        
        return true
        
    } catch Error as e {
        if (progress && progress.gui) {
            progress.gui.Destroy()
        }
        MsgBox("Installation failed: " e.Message, "Installation Error", 16)
        return false
    }
}

/**
 * Initialize installation system and execute setup if required
 * Sets up global paths and handles first-time installation or upgrades
 * @return {Void}
 */
InitializeInstallation() {
    global INSTALL_DIR, SETTINGS_FILE, FIRST_RUN
    
    ; Set up installation paths
    INSTALL_DIR := GetInstallationPath()
    SETTINGS_FILE := INSTALL_DIR "\settings.ini"
    
    ; Analyze current installation status
    installStatus := CheckInstallationStatus()
    
    if (installStatus.isFirstRun) {
        FIRST_RUN := true
        
        if (!PerformInstallation()) {
            MsgBox("Installation failed. The program will exit.", "Setup Error", 16)
            ExitApp()
        }
    } else if (installStatus.needsUpgrade) {
        FIRST_RUN := false
        
        if (AskForUpgrade(SCRIPT_VERSION, installStatus.installedVersion)) {
            if (PerformUpgrade(installStatus.installedVersion)) {
                if (A_ScriptDir = INSTALL_DIR) {
                    LoadUserSettings()
                } else {
                    ; Restart from upgraded installation
                    newExecutable := INSTALL_DIR "\" A_ScriptName
                    if (FileExist(newExecutable)) {
                        Run(newExecutable)
                        ExitApp()
                    }
                }
            } else {
                LoadUserSettings()
            }
        } else {
            LoadUserSettings()
        }
    } else {
        FIRST_RUN := false
        LoadUserSettings()
    }
}

/**
 * Load user settings from configuration file into global variables
 * @return {Void}
 */
LoadUserSettings() {
    global SETTINGS_FILE, quietMode, debugMode, captureHotkey, searchHotkey
    
    try {
        quietMode := (IniRead(SETTINGS_FILE, "General", "QuietMode", "false") = "true")
        debugMode := (IniRead(SETTINGS_FILE, "General", "DebugMode", "false") = "true")
        
        captureHotkey := IniRead(SETTINGS_FILE, "Hotkeys", "CaptureWindow", "`` & 2")
        searchHotkey := IniRead(SETTINGS_FILE, "Hotkeys", "SearchShortcut", "`` & 3")

    } catch 
{
        ; Use default values if settings cannot be read
        quietMode := false
        debugMode := false
        captureHotkey := "`` & 2"
        searchHotkey := "`` & 3"
    }
}

;==============================================================================
; UNINSTALLATION SYSTEM
;==============================================================================

/**
 * Display uninstall confirmation dialog with data preservation options
 * @return {Object} User choice: {uninstall: Boolean, keepData: Boolean}
 */
ShowUninstallDialog() {
    global SCRIPT_NAME, MyGui
    
    ; COMPREHENSIVE THEME CHECK BEFORE DIALOG CREATION
    ForceThemeCheck()
    
    try {
        UninstallGui := CreateManagedDialog("+Owner" MyGui.Hwnd, "Uninstall " SCRIPT_NAME)
    } catch {
        UninstallGui := Gui("+Owner" MyGui.Hwnd, "Uninstall " SCRIPT_NAME)
    }
    UninstallGui.SetFont("s10")
    ApplyDarkModeToGUI(UninstallGui, "Uninstall Dialog")
    colors := GetThemeColors()
    isDark := IsWindowsDarkMode()
    
    ; Dialog header
    headerText := UninstallGui.Add("Text", "x10 y10 w350 center", "⚠️  Uninstall Confirmation")
    headerText.SetFont("bold s11")
    ApplyControlTheming(headerText, "Text", colors, isDark)
    
    ; Main confirmation text
    UninstallGui.SetFont("s10")
    confirmText := UninstallGui.Add("Text", "x10 y40 w350", "Are you sure you want to uninstall " SCRIPT_NAME "?")
    ApplyControlTheming(confirmText, "Text", colors, isDark)
    
    ; Items to be removed list
    removeText := UninstallGui.Add("Text", "x10 y70 w350", "This will remove:")
    ApplyControlTheming(removeText, "Text", colors, isDark)
    
    item1 := UninstallGui.Add("Text", "x20 y90 w330", "• The application executable and files")
    ApplyControlTheming(item1, "Text", colors, isDark)
    
    item2 := UninstallGui.Add("Text", "x20 y110 w330", "• Windows startup integration")
    ApplyControlTheming(item2, "Text", colors, isDark)
    
    item3 := UninstallGui.Add("Text", "x20 y130 w330", "• Desktop shortcut")
    ApplyControlTheming(item3, "Text", colors, isDark)
    
    item4 := UninstallGui.Add("Text", "x20 y150 w330", "• Installation directory")
    ApplyControlTheming(item4, "Text", colors, isDark)
    
    ; User data options
    DataGroupBox := UninstallGui.Add("GroupBox", "x10 y180 w350 h80", "User Data")
    ApplyControlTheming(DataGroupBox, "GroupBox", colors, isDark)
    
    KeepDataRadio := UninstallGui.Add("Radio", "x20 y205 w320 Checked", "&Keep my shortcuts database and settings")
    ApplyControlTheming(KeepDataRadio, "Button", colors, isDark)
    
    keepHelpText := UninstallGui.Add("Text", "x40 y225 w300", "Preserve your shortcuts and preferences for future reinstallation")
    ApplyControlTheming(keepHelpText, "Text", colors, isDark)
    
    DeleteAllRadio := UninstallGui.Add("Radio", "x20 y245 w320", "&Delete everything (complete removal)")
    ApplyControlTheming(DeleteAllRadio, "Button", colors, isDark)
    
    ; Action buttons
    UninstallButton := UninstallGui.Add("Button", "x70 y280 w100", "&Uninstall")
    ApplyControlTheming(UninstallButton, "Button", colors, isDark)
    
    CancelUninstallButton := UninstallGui.Add("Button", "Default x180 y280 w100", "&Cancel")
    ApplyControlTheming(CancelUninstallButton, "Button", colors, isDark)
    
    userChoice := {uninstall: false, keepData: true}
    
    ConfirmUninstall(*) {
        keepData := KeepDataRadio.Value
        
        confirmMsg := "This action cannot be undone.`n`n"
        if (keepData) {
            confirmMsg .= "Your shortcuts database and settings will be preserved."
        } else {
            confirmMsg .= "ALL data including your shortcuts database will be permanently deleted."
        }
        confirmMsg .= "`n`nContinue with uninstall?"
        
        result := MsgBox(confirmMsg, "Final Confirmation", 52)
        
        if (result = "Yes") {
            userChoice.uninstall := true
            userChoice.keepData := keepData
            UninstallGui.Destroy()
        }
    }
    
    UninstallButton.OnEvent("Click", ConfirmUninstall)
    CancelUninstallButton.OnEvent("Click", (*) => UninstallGui.Destroy())
    UninstallGui.OnEvent("Close", (*) => UninstallGui.Destroy())
    
    RegisterGuiForEscapeHandling(UninstallGui, (*) => UninstallGui.Destroy())
    
    UninstallGui.Show("w370 h320")
    ForceWindowToForeground(UninstallGui.Hwnd)
    
    ; Wait for dialog completion
    WinWaitClose("ahk_id " UninstallGui.Hwnd)
    
    return userChoice
}

/**
 * Execute complete uninstallation process with optional data preservation
 * @param {Boolean} keepUserData - True to preserve user data in backup location
 * @return {Boolean} True if uninstallation completed successfully
 */
PerformUninstall(keepUserData := true) {
    global INSTALL_DIR, SCRIPT_NAME, SETTINGS_FILE, DB_FILE
    
    ; Create progress dialog
    try {
        ProgressGui := Gui("+ToolWindow -MaximizeBox -MinimizeBox", "Uninstalling " SCRIPT_NAME)
        ProgressGui.SetFont("s10")
        ApplyDarkModeToGUI(ProgressGui, "Uninstall Progress")
        colors := GetThemeColors()
        isDark := IsWindowsDarkMode()
        
        titleText := ProgressGui.Add("Text", "x20 y20 w360 center", "Uninstalling " SCRIPT_NAME "...")
        ApplyControlTheming(titleText, "Text", colors, isDark)
        
        ProgressBar := ProgressGui.Add("Progress", "x20 y50 w360 h20", 0)
        
        StatusText := ProgressGui.Add("Text", "x20 y80 w360 center", "Starting uninstall...")
        ApplyControlTheming(StatusText, "Text", colors, isDark)
        
        ProgressGui.Show("w400 h120")
    } catch {
        ProgressGui := false
    }
    
    try {
        ; Phase 1: Remove Windows integration
        if (ProgressGui) {
            ProgressBar.Value := 20
            StatusText.Text := "Removing startup integration..."
            Sleep(500)
        }
        
        try {
            regKey := "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
            RegDelete(regKey, SCRIPT_NAME)
        } catch {
            ; Registry key might not exist
        }
        
        ; Phase 2: Remove desktop shortcut
        if (ProgressGui) {
            ProgressBar.Value := 40
            StatusText.Text := "Removing desktop shortcut..."
            Sleep(500)
        }
        
        try {
            desktopPath := EnvGet("USERPROFILE") "\Desktop"
            shortcutPath := desktopPath "\" SCRIPT_NAME ".lnk"
            if (FileExist(shortcutPath)) {
                FileDelete(shortcutPath)
            }
        } catch {
            ; Continue if shortcut removal fails
        }
        
        ; Phase 3: Handle user data preservation
        if (ProgressGui) {
            ProgressBar.Value := 60
            StatusText.Text := keepUserData ? "Preserving user data..." : "Removing user data..."
            Sleep(500)
        }
        
        if (keepUserData) {
            backupDir := EnvGet("USERPROFILE") "\Documents\" SCRIPT_NAME "_UserData"
            
            try {
                if (!DirExist(backupDir)) {
                    DirCreate(backupDir)
                }
                
                ; Backup configuration and database files
                if (FileExist(SETTINGS_FILE)) {
                    FileCopy(SETTINGS_FILE, backupDir "\settings.ini", 1)
                }
                if (FileExist(DB_FILE)) {
                    FileCopy(DB_FILE, backupDir "\appinfo.db", 1)
                }
                if (DirExist(INSTALL_DIR "\Backups")) {
                    DirCopy(INSTALL_DIR "\Backups", backupDir "\Backups", 1)
                }
                if (DirExist(INSTALL_DIR "\Exports")) {
                    DirCopy(INSTALL_DIR "\Exports", backupDir "\Exports", 1)
                }
            } catch {
                ; Continue even if backup fails
            }
        }
        
        ; Phase 4: Remove installation directory
        if (ProgressGui) {
            ProgressBar.Value := 80
            StatusText.Text := "Removing installation files..."
            Sleep(500)
        }
        
        Sleep(1000) ; Allow file handles to close
        
        try {
            if (DirExist(INSTALL_DIR)) {
                DirDelete(INSTALL_DIR, 1)
            }
        } catch Error as e {
            ; Try to delete individual files if directory deletion fails
            try {
                if (FileExist(DB_FILE))
                    FileDelete(DB_FILE)
                if (FileExist(SETTINGS_FILE))
                    FileDelete(SETTINGS_FILE)
                    
                if (DirExist(INSTALL_DIR "\Backups"))
                    DirDelete(INSTALL_DIR "\Backups", 1)
                if (DirExist(INSTALL_DIR "\Exports"))
                    DirDelete(INSTALL_DIR "\Exports", 1)
                if (DirExist(INSTALL_DIR "\Logs"))
                    DirDelete(INSTALL_DIR "\Logs", 1)
            } catch {
                ; Some files might be in use
            }
        }
        
        ; Phase 5: Completion
        if (ProgressGui) {
            ProgressBar.Value := 100
            StatusText.Text := "Uninstall complete!"
            Sleep(1000)
            ProgressGui.Destroy()
        }
        
        ; Display completion message
        completionMsg := SCRIPT_NAME " has been successfully uninstalled.`n`n"
        
        if (keepUserData) {
            completionMsg .= "Your shortcuts database and settings have been preserved in:`n"
            completionMsg .= EnvGet("USERPROFILE") "\Documents\" SCRIPT_NAME "_UserData`n`n"
            completionMsg .= "You can reinstall later and your data will be restored."
        } else {
            completionMsg .= "All files and data have been completely removed."
        }
        
        MsgBox(completionMsg, "Uninstall Complete", 64)
        return true
        
    } catch Error as e {
        if (ProgressGui) {
            ProgressGui.Destroy()
        }
        MsgBox("Uninstall encountered an error: " e.Message, "Uninstall Error", 48)
        return false
    }
}

/**
 * Handle uninstall process initiation with safety checks
 * @return {Void}
 */
HandleUninstall(*) {
    global INSTALL_DIR, SCRIPT_NAME
    
    ; Verify uninstall can only be performed from installed version
    if (A_ScriptDir != INSTALL_DIR) {
        MsgBox("Uninstall can only be performed from the installed version.`n`n" 
             . "Please run the application from:`n" INSTALL_DIR, "Uninstall Error", 48)
        return
    }
    
    userChoice := ShowUninstallDialog()
    
    if (userChoice.uninstall) {
        if (PerformUninstall(userChoice.keepData)) {
            ExitApp()
        }
    }
}

;==============================================================================
; SQLITE DATABASE INTERFACE
;==============================================================================

/**
 * Check if SQLite library is available for database operations
 * @return {Boolean} True if SQLite library can be loaded and used
 */
SQLiteAvailable() {
    global INSTALL_DIR
    
    ; Check installation directory first (permanent location)
    if (INSTALL_DIR && INSTALL_DIR != "") {
        dllPath := INSTALL_DIR "\sqlite3.dll"
        if FileExist(dllPath) {
            return true
        }
    }
    
    ; Check temp directory (where FileInstall extracts for compiled versions)
    tempDir := A_Temp "\Keyfinder_Temp"
    dllPath := tempDir "\sqlite3.dll"
    if FileExist(dllPath) {
        return true
    }
    
    ; Check script directory (for non-compiled versions)
    dllPath := A_ScriptDir "\sqlite3.dll"
    if FileExist(dllPath) {
        return true
    }
    
    return false
}

/**
 * Initialize SQLite system with error handling
 * @return {Boolean} True if SQLite is ready for database operations
 */
InitializeSQLite() {
    if (!SQLiteAvailable()) {
        if A_IsCompiled {
            MsgBox("SQLite database support is not available.`n`n" 
                 . "The application will use backup files instead.", "Database Warning", 48)
        }
        return false
    }
    return true
}

/**
 * Open SQLite database connection with enhanced error handling and retry logic
 * @param {String} dbFile - Full path to database file
 * @return {Integer|Boolean} Database handle on success, false on failure
 */
SQLite_Open(dbFile) {
    global INSTALL_DIR
    static SQLITE_OPEN_READWRITE := 0x2
    static SQLITE_OPEN_CREATE := 0x4
    
    ; Test if database file is accessible (not locked by another process)
    if FileExist(dbFile) {
        testFile := dbFile . ".test"
        try {
            FileMove(dbFile, testFile)
            FileMove(testFile, dbFile)
        } catch {
            return false
        }
    }
    
    ; Load SQLite library
    ; Try from INSTALL_DIR first (permanent location)
    hSQLiteModule := 0
    if (INSTALL_DIR && INSTALL_DIR != "") {
        dllPath := INSTALL_DIR "\sqlite3.dll"
        if FileExist(dllPath) {
            hSQLiteModule := DllCall("LoadLibrary", "Str", dllPath, "Ptr")
        }
    }
    
    ; Fall back to temp directory if INSTALL_DIR load failed
    if (!hSQLiteModule) {
        tempDir := A_Temp "\Keyfinder_Temp"
        dllPath := tempDir "\sqlite3.dll"
        if FileExist(dllPath) {
            hSQLiteModule := DllCall("LoadLibrary", "Str", dllPath, "Ptr")
        }
    }
    
    ; Final fallback to script directory
    if (!hSQLiteModule) {
        hSQLiteModule := DllCall("LoadLibrary", "Str", "sqlite3.dll", "Ptr")
    }
    
    if (!hSQLiteModule) {
        return false
    }
    
    ; Attempt to open database with retry mechanism
    maxRetries := 3
    retryCount := 0
    
    while (retryCount < maxRetries) {
        result := DllCall("sqlite3\sqlite3_open_v2", "AStr", dbFile, "Ptr*", &hDB := 0, "Int", SQLITE_OPEN_READWRITE | SQLITE_OPEN_CREATE, "Ptr", 0, "Int")
        
        if (result = 0) {
            ; Set busy timeout for concurrent access
            DllCall("sqlite3\sqlite3_busy_timeout", "Ptr", hDB, "Int", 5000)
            return hDB
        } else if (result = 5) {  ; SQLITE_BUSY
            retryCount++
            Sleep(1000)
        } else {
            ; Other error occurred
            if (hDB)
                DllCall("sqlite3\sqlite3_close", "Ptr", hDB)
            return false
        }
    }
    
    ; All retries failed
    if (hDB)
        DllCall("sqlite3\sqlite3_close", "Ptr", hDB)
    return false
}

/**
 * Close SQLite database connection safely
 * @param {Integer} hDB - Database handle from SQLite_Open
 * @return {Void}
 */
SQLite_Close(hDB) {
    if hDB
        DllCall("sqlite3\sqlite3_close", "Ptr", hDB)
}

/**
 * Execute SQL statement without expecting return data
 * @param {Integer} hDB - Database handle
 * @param {String} SQL - SQL statement to execute
 * @return {Boolean} True if execution successful
 * @throws {Error} On SQL execution failure
 */
SQLite_Exec(hDB, SQL) {
    result := DllCall("sqlite3\sqlite3_exec", "Ptr", hDB, "AStr", SQL, "Ptr", 0, "Ptr", 0, "Ptr*", &errMsg := 0, "Int")
    
    if result != 0 {
        sqlError := StrGet(errMsg, "UTF-8")
        DllCall("sqlite3\sqlite3_free", "Ptr", errMsg)
        throw Error("SQLite error: " sqlError)
    }
    
    return true
}

/**
 * Execute SQL query and return results as table array
 * @param {Integer} hDB - Database handle
 * @param {String} SQL - SQL query to execute
 * @return {Array} Multi-dimensional array: [headers][row1][row2]...
 * @throws {Error} On SQL execution failure
 */
SQLite_GetTable(hDB, SQL) {
    result := DllCall("sqlite3\sqlite3_get_table", "Ptr", hDB, "AStr", SQL, "Ptr*", &table := 0, "Int*", &rows := 0, "Int*", &cols := 0, "Ptr*", &errMsg := 0, "Int")
    
    if result != 0 {
        sqlError := StrGet(errMsg, "UTF-8")
        DllCall("sqlite3\sqlite3_free", "Ptr", errMsg)
        throw Error("SQLite error: " sqlError)
    }
    
    ; Build result array with proper structure
    resultArray := Array()
    resultArray.Length := rows + 1
    
    ; First row contains column headers
    Loop cols {
        if A_Index = 1
            resultArray[1] := Array()
        colHeader := StrGet(NumGet(table, (A_Index - 1) * A_PtrSize, "Ptr"), "UTF-8")
        resultArray[1].Push(colHeader)
    }
    
    ; Subsequent rows contain data
    Loop rows {
        row := A_Index + 1
        resultArray[row] := Array()
        
        Loop cols {
            col := A_Index
            value := StrGet(NumGet(table, ((row - 1) * cols + col - 1) * A_PtrSize, "Ptr"), "UTF-8")
            resultArray[row].Push(value)
        }
    }
    
    ; Free SQLite-allocated memory
    DllCall("sqlite3\sqlite3_free_table", "Ptr", table)
    return resultArray
}

/**
 * Get last SQLite error message for debugging
 * @param {Integer} hDB - Database handle (optional)
 * @return {String} Human-readable error message
 */
SQLite_LastError(hDB := 0) {
    return StrGet(DllCall("sqlite3\sqlite3_errmsg", "Ptr", hDB, "Ptr"), "UTF-8")
}

/**
 * Execute parameterized SQL statement (INSERT, UPDATE, DELETE)
 * @param {Integer} hDB - Database handle
 * @param {String} SQL - SQL with parameter placeholders (?)
 * @param {Any*} params - Parameters to bind to placeholders
 * @return {Boolean} True if execution successful
 */
SQLite_PrepareBindExecute(hDB, SQL, params*) {
    ; Prepare SQL statement
    result := DllCall("sqlite3\sqlite3_prepare_v2", "Ptr", hDB, "AStr", SQL, "Int", -1, "Ptr*", &stmt := 0, "Ptr", 0, "Int")
    if (result != 0) {
        return false
    }
    
    ; Bind parameters to placeholders
    for i, param in params {
        paramIndex := i
        
        if IsInteger(param)
            DllCall("sqlite3\sqlite3_bind_int", "Ptr", stmt, "Int", paramIndex, "Int", param, "Int")
        else if IsFloat(param)
            DllCall("sqlite3\sqlite3_bind_double", "Ptr", stmt, "Int", paramIndex, "Double", param, "Int")
        else if IsObject(param)
            DllCall("sqlite3\sqlite3_bind_null", "Ptr", stmt, "Int", paramIndex, "Int")
        else
            DllCall("sqlite3\sqlite3_bind_text", "Ptr", stmt, "Int", paramIndex, "AStr", param, "Int", -1, "Ptr", -1, "Int")
    }
    
    ; Execute prepared statement
    result := DllCall("sqlite3\sqlite3_step", "Ptr", stmt, "Int")
    success := (result = 101) ; SQLITE_DONE
    
    ; Clean up statement resources
    DllCall("sqlite3\sqlite3_finalize", "Ptr", stmt, "Int")
    return success
}

/**
 * Execute parameterized SQL query and return results as object array
 * @param {Integer} hDB - Database handle
 * @param {String} SQL - SQL query with parameter placeholders (?)
 * @param {Any*} params - Parameters to bind to placeholders
 * @return {Array} Array of result objects with column names as properties
 */
SQLite_Query(hDB, SQL, params*) {
    ; Prepare SQL statement
    result := DllCall("sqlite3\sqlite3_prepare_v2", "Ptr", hDB, "AStr", SQL, "Int", -1, "Ptr*", &stmt := 0, "Ptr", 0, "Int")
    if (result != 0) {
        return []
    }
    
    ; Bind parameters to placeholders
    for i, param in params {
        paramIndex := i
        
        if IsInteger(param)
            DllCall("sqlite3\sqlite3_bind_int", "Ptr", stmt, "Int", paramIndex, "Int", param, "Int")
        else if IsFloat(param)
            DllCall("sqlite3\sqlite3_bind_double", "Ptr", stmt, "Int", paramIndex, "Double", param, "Int")
        else if IsObject(param)
            DllCall("sqlite3\sqlite3_bind_null", "Ptr", stmt, "Int", paramIndex, "Int")
        else
            DllCall("sqlite3\sqlite3_bind_text", "Ptr", stmt, "Int", paramIndex, "AStr", param, "Int", -1, "Ptr", -1, "Int")
    }
    
    results := []
    
    ; Process result rows
    while (DllCall("sqlite3\sqlite3_step", "Ptr", stmt, "Int") = 100) { ; SQLITE_ROW
        row := {}
        colCount := DllCall("sqlite3\sqlite3_column_count", "Ptr", stmt, "Int")
        
        Loop colCount {
            colIndex := A_Index - 1
            colName := StrGet(DllCall("sqlite3\sqlite3_column_name", "Ptr", stmt, "Int", colIndex, "Ptr"), "UTF-8")
            colType := DllCall("sqlite3\sqlite3_column_type", "Ptr", stmt, "Int", colIndex, "Int")
            
            ; Extract value based on SQLite data type
            switch colType {
                case 1: ; SQLITE_INTEGER
                    row[colName] := DllCall("sqlite3\sqlite3_column_int", "Ptr", stmt, "Int", colIndex, "Int")
                case 2: ; SQLITE_FLOAT
                    row[colName] := DllCall("sqlite3\sqlite3_column_double", "Ptr", stmt, "Int", colIndex, "Double")
                case 3: ; SQLITE_TEXT
                    row[colName] := StrGet(DllCall("sqlite3\sqlite3_column_text", "Ptr", stmt, "Int", colIndex, "Ptr"), "UTF-8")
                case 4: ; SQLITE_BLOB
                    row[colName] := "[BLOB]"
                case 5: ; SQLITE_NULL
                    row[colName] := ""
            }
        }
        
        results.Push(row)
    }
    
    ; Clean up statement resources
    DllCall("sqlite3\sqlite3_finalize", "Ptr", stmt, "Int")
    return results
}

;==============================================================================
; DATABASE MANAGEMENT SYSTEM
;==============================================================================

/**
 * Create empty database with proper table structure and indexes
 * @return {Void}
 */
CreateEmptyDatabase() {
    global DB_FILE
    
    try {
        db := SQLite_Open(DB_FILE)
        if (db) {
            ; Create main tables with proper schema
            SQLite_Exec(db, "CREATE TABLE shortcuts (id INTEGER PRIMARY KEY, process_name TEXT, command_name TEXT, shortcut_key TEXT, category TEXT, description TEXT);")
            SQLite_Exec(db, "CREATE TABLE at_compatibility (id INTEGER PRIMARY KEY, process_name TEXT, at_software TEXT, compatibility_level TEXT, known_issues TEXT, workarounds TEXT);")
            SQLite_Exec(db, "CREATE TABLE at_usage (id INTEGER PRIMARY KEY, process_name TEXT, window_class TEXT, at_software TEXT, timestamp TEXT);")
            
            ; Create indexes for performance
            SQLite_Exec(db, "CREATE INDEX idx_shortcuts_process ON shortcuts(process_name);")
            SQLite_Exec(db, "CREATE INDEX idx_at_compat_process ON at_compatibility(process_name);")
            
            SQLite_Close(db)
        }
    } catch Error as e {
        ; Fall back to file-based storage if database creation fails
    }
}

/**
 * Load data from SQLite database into memory structures
 * @param {Integer} db - Database handle
 * @return {Void}
 */
LoadDataFromDatabase(db) {
    global shortcuts_db, at_compat_db
    
    try {
        ; Load shortcuts data organized by process name
        processNamesTable := SQLite_GetTable(db, "SELECT DISTINCT process_name FROM shortcuts;")
        
        if (processNamesTable.Length > 1) {
            Loop processNamesTable.Length - 1 {
                rowIndex := A_Index + 1
                
                if (processNamesTable[rowIndex].Length > 0) {
                    processName := processNamesTable[rowIndex][1]
                    
                    if (!processName)
                        continue
                    
                    ; Load shortcuts for this process
                    shortcutsTable := SQLite_GetTable(db, "SELECT command_name, shortcut_key, category, description FROM shortcuts WHERE process_name = '" processName "';")
                    
                    if (shortcutsTable.Length > 1) {
                        shortcuts := []
                        
                        Loop shortcutsTable.Length - 1 {
                            rowIdx := A_Index + 1
                            
                            if (shortcutsTable[rowIdx].Length >= 4) {
                                shortcuts.Push({
                                    command_name: shortcutsTable[rowIdx][1],
                                    shortcut_key: shortcutsTable[rowIdx][2],
                                    category: shortcutsTable[rowIdx][3] ? shortcutsTable[rowIdx][3] : "General",
                                    description: shortcutsTable[rowIdx][4] ? shortcutsTable[rowIdx][4] : ""
                                })
                            }
                        }
                        
                        if (shortcuts.Length > 0) {
                            shortcuts_db[processName] := shortcuts
                        }
                    }
                }
            }
        }
        
        ; Load accessibility compatibility data
        compatTable := SQLite_GetTable(db, "SELECT process_name, at_software, compatibility_level, known_issues, workarounds FROM at_compatibility;")
        
        if (compatTable.Length > 1) {
            compatMap := Map()
            
            Loop compatTable.Length - 1 {
                rowIdx := A_Index + 1
                
                if (compatTable[rowIdx].Length >= 5) {
                    processName := compatTable[rowIdx][1]
                    
                    if (!compatMap.Has(processName))
                        compatMap[processName] := []
                    
                    compatMap[processName].Push({
                        at_software: compatTable[rowIdx][2],
                        compatibility_level: compatTable[rowIdx][3],
                        known_issues: compatTable[rowIdx][4],
                        workarounds: compatTable[rowIdx][5]
                    })
                }
            }
            
            at_compat_db := compatMap
        }
        
    } catch Error as e {
        ; Continue with empty data structures if loading fails
    }
}

/**
 * Populate database with initial example data for demonstration
 * @return {Void}
 */
PopulateInitialData() {
    global DB_FILE
    
    if (!InitializeSQLite())
        return
        
    db := SQLite_Open(DB_FILE)
    if (!db)
        return
    
    try {
        ; Add comprehensive Chrome shortcuts
        AddShortcutToDatabase(db, "chrome.exe", "New Tab", "Ctrl+T", "Navigation", "Opens a new browser tab")
        AddShortcutToDatabase(db, "chrome.exe", "New Window", "Ctrl+N", "Navigation", "Opens a new browser window")
        AddShortcutToDatabase(db, "chrome.exe", "Close Tab", "Ctrl+W", "Navigation", "Closes the current tab")
        AddShortcutToDatabase(db, "chrome.exe", "Reopen Closed Tab", "Ctrl+Shift+T", "Navigation", "Reopens the last closed tab")
        AddShortcutToDatabase(db, "chrome.exe", "Find", "Ctrl+F", "Navigation", "Find text on page")
        AddShortcutToDatabase(db, "chrome.exe", "Refresh", "F5", "Navigation", "Refresh the current page")
        AddShortcutToDatabase(db, "chrome.exe", "Bookmarks", "Ctrl+Shift+B", "Navigation", "Show/hide bookmarks bar")
        
        ; Add Notepad shortcuts
        AddShortcutToDatabase(db, "notepad.exe", "New", "Ctrl+N", "File", "Creates a new document")
        AddShortcutToDatabase(db, "notepad.exe", "Open", "Ctrl+O", "File", "Opens an existing document")
        AddShortcutToDatabase(db, "notepad.exe", "Save", "Ctrl+S", "File", "Saves the current document")
        AddShortcutToDatabase(db, "notepad.exe", "Save As", "Ctrl+Shift+S", "File", "Save document with new name")
        AddShortcutToDatabase(db, "notepad.exe", "Find", "Ctrl+F", "Edit", "Opens the find dialog")
        AddShortcutToDatabase(db, "notepad.exe", "Replace", "Ctrl+H", "Edit", "Opens find and replace dialog")
        AddShortcutToDatabase(db, "notepad.exe", "Select All", "Ctrl+A", "Edit", "Select all text")
        
        ; Add Microsoft Word shortcuts
        AddShortcutToDatabase(db, "winword.exe", "Bold", "Ctrl+B", "Text Formatting", "Makes selected text bold")
        AddShortcutToDatabase(db, "winword.exe", "Italic", "Ctrl+I", "Text Formatting", "Makes selected text italic")
        AddShortcutToDatabase(db, "winword.exe", "Underline", "Ctrl+U", "Text Formatting", "Underlines selected text")
        AddShortcutToDatabase(db, "winword.exe", "Save", "Ctrl+S", "File Operations", "Saves the current document")
        AddShortcutToDatabase(db, "winword.exe", "New Document", "Ctrl+N", "File Operations", "Creates a new document")
        AddShortcutToDatabase(db, "winword.exe", "Print", "Ctrl+P", "File Operations", "Opens print dialog")
        AddShortcutToDatabase(db, "winword.exe", "Find", "Ctrl+F", "Navigation", "Opens find dialog")
        AddShortcutToDatabase(db, "winword.exe", "Cut", "Ctrl+X", "Edit", "Cut selected text")
        AddShortcutToDatabase(db, "winword.exe", "Copy", "Ctrl+C", "Edit", "Copy selected text")
        AddShortcutToDatabase(db, "winword.exe", "Paste", "Ctrl+V", "Edit", "Paste from clipboard")
        
        ; Add Excel shortcuts
        AddShortcutToDatabase(db, "excel.exe", "New Workbook", "Ctrl+N", "File Operations", "Creates a new workbook")
        AddShortcutToDatabase(db, "excel.exe", "Save", "Ctrl+S", "File Operations", "Saves the current workbook")
        AddShortcutToDatabase(db, "excel.exe", "Bold", "Ctrl+B", "Text Formatting", "Makes selected cells bold")
        AddShortcutToDatabase(db, "excel.exe", "Insert Row", "Ctrl+Shift++", "Edit", "Insert a new row")
        AddShortcutToDatabase(db, "excel.exe", "Delete Row", "Ctrl+-", "Edit", "Delete selected row")
        AddShortcutToDatabase(db, "excel.exe", "AutoSum", "Alt+=", "Functions", "Insert AutoSum formula")
        
        ; Add Windows Explorer shortcuts
        AddShortcutToDatabase(db, "explorer.exe", "New Folder", "Ctrl+Shift+N", "File Operations", "Create a new folder")
        AddShortcutToDatabase(db, "explorer.exe", "Copy", "Ctrl+C", "File Operations", "Copy selected items")
        AddShortcutToDatabase(db, "explorer.exe", "Cut", "Ctrl+X", "File Operations", "Cut selected items")
        AddShortcutToDatabase(db, "explorer.exe", "Paste", "Ctrl+V", "File Operations", "Paste items from clipboard")
        AddShortcutToDatabase(db, "explorer.exe", "Select All", "Ctrl+A", "Edit", "Select all items")
        AddShortcutToDatabase(db, "explorer.exe", "Properties", "Alt+Enter", "View", "View properties of selected item")
        
        ; Add accessibility technology compatibility data
        SQLite_PrepareBindExecute(db, 
            "INSERT INTO at_compatibility (process_name, at_software, compatibility_level, known_issues, workarounds) VALUES (?, ?, ?, ?, ?)",
            "chrome.exe", "NVDA Screen Reader", "Good", "Some dynamic content may not be announced automatically", "Use NVDA's browse mode and element navigation")
            
        SQLite_PrepareBindExecute(db, 
            "INSERT INTO at_compatibility (process_name, at_software, compatibility_level, known_issues, workarounds) VALUES (?, ?, ?, ?, ?)",
            "chrome.exe", "JAWS Screen Reader", "Good", "Custom controls sometimes lack proper labels", "Use JAWS cursor for unlabeled elements")
            
        SQLite_PrepareBindExecute(db, 
            "INSERT INTO at_compatibility (process_name, at_software, compatibility_level, known_issues, workarounds) VALUES (?, ?, ?, ?, ?)",
            "winword.exe", "NVDA Screen Reader", "Excellent", "Complex layout may be difficult to navigate", "Use document navigation shortcuts (Ctrl+Arrow keys)")
            
        SQLite_PrepareBindExecute(db, 
            "INSERT INTO at_compatibility (process_name, at_software, compatibility_level, known_issues, workarounds) VALUES (?, ?, ?, ?, ?)",
            "winword.exe", "JAWS Screen Reader", "Excellent", "Table navigation can be complex", "Use JAWS table navigation commands")
            
    } catch Error as e {
        ; Error adding initial data - continue with empty database
    }
    
    SQLite_Close(db)
}

/**
 * Add single shortcut entry to database using parameterized query
 * @param {Integer} db - Database handle
 * @param {String} processName - Process executable name
 * @param {String} commandName - Human-readable command name
 * @param {String} shortcutKey - Keyboard shortcut combination
 * @param {String} category - Shortcut category (default "General")
 * @param {String} description - Detailed description (default empty)
 * @return {Boolean} True if insertion successful
 */
AddShortcutToDatabase(db, processName, commandName, shortcutKey, category := "General", description := "") {
    if (!db)
        return false

    sql := "INSERT INTO shortcuts (process_name, command_name, shortcut_key, category, description) VALUES (?, ?, ?, ?, ?)"
    return SQLite_PrepareBindExecute(db, sql, processName, commandName, shortcutKey, category, description)
}

/**
 * Load shortcuts from backup INI file when SQLite is unavailable
 * @return {Boolean} True if shortcuts were successfully loaded from file
 */
LoadShortcutsFromFile() {
    global shortcuts_db, INSTALL_DIR
    
    ; Try multiple backup file locations
    backupFile := INSTALL_DIR "\Backups\shortcuts_backup.ini"
    if !FileExist(backupFile) {
        backupFile := A_ScriptDir "\shortcuts_backup.ini"
        if !FileExist(backupFile) {
            return false
        }
    }
        
    try {
        fileContent := FileRead(backupFile)
        if (fileContent = "") {
            return false
        }
        
        ; Parse INI-style backup file
        currentApp := ""
        shortcuts := []
        tempShortcut := Map()
        shortcutCount := 0
        
        lines := StrSplit(fileContent, "`n", "`r")
        
        for line in lines {
            if (line = "")
                continue
            
            ; Process section headers [application.exe]
            if (SubStr(line, 1, 1) = "[" && SubStr(line, -1) = "]") {
                ; Save previous application's shortcuts
                if (currentApp != "" && shortcuts.Length > 0) {
                    shortcuts_db[currentApp] := shortcuts.Clone()
                }
                
                currentApp := SubStr(line, 2, StrLen(line) - 2)
                shortcuts := []
                continue
            }
            
            ; Process key-value pairs
            if (InStr(line, "=")) {
                parts := StrSplit(line, "=", "", 2)
                prop := parts[1]
                value := parts[2]
                
                ; Parse shortcut properties (shortcut1_category=Navigation)
                if (InStr(prop, "shortcut") && InStr(prop, "_")) {
                    propParts := StrSplit(prop, "_", "", 2)
                    shortcutNumStr := SubStr(propParts[1], 9)  ; Remove "shortcut" prefix
                    propName := propParts[2]
                    
                    shortcutNum := Integer(shortcutNumStr)
                    shortcutKey := "s" shortcutNumStr
                    
                    if (!tempShortcut.Has(shortcutKey))
                        tempShortcut[shortcutKey] := Map()
                    
                    ; Map property names to standardized keys
                    if (propName = "category")
                        tempShortcut[shortcutKey]["category"] := value
                    else if (propName = "command")
                        tempShortcut[shortcutKey]["command_name"] := value
                    else if (propName = "keys")
                        tempShortcut[shortcutKey]["shortcut_key"] := value
                    else if (propName = "desc")
                        tempShortcut[shortcutKey]["description"] := value
                    
                    shortcutCount := Max(shortcutCount, shortcutNum)
                }
            }
        }
        
        ; Process final application
        if (currentApp != "") {
            shortcuts := []
            
            Loop shortcutCount {
                i := A_Index
                shortcutKey := "s" i
                if (tempShortcut.Has(shortcutKey)) {
                    shortcutData := tempShortcut[shortcutKey]
                    
                    if (shortcutData.Count > 0) {
                        if (shortcutData.Has("command_name") && shortcutData.Has("shortcut_key") && 
                            shortcutData["command_name"] != "" && shortcutData["shortcut_key"] != "") {
                            
                            shortcutObj := {
                                category: shortcutData.Get("category", "General"),
                                command_name: shortcutData["command_name"],
                                shortcut_key: shortcutData["shortcut_key"],
                                description: shortcutData.Get("description", "")
                            }
                            shortcuts.Push(shortcutObj)
                        }
                    }
                }
            }
            
            if (shortcuts.Length > 0)
                shortcuts_db[currentApp] := shortcuts
        }
        
        return shortcuts_db.Count > 0
    } catch Error as e {
        MsgBox("Error loading shortcuts from backup file: " e.Message, "Load Error", 16)
        return false
    }
}

/**
 * Create example shortcuts data when no database is available
 * @return {Void}
 */
AddExampleShortcuts() {
    global shortcuts_db
    
    ; Chrome browser shortcuts
    chrome_shortcuts := [
        {category: "Navigation", command_name: "New Tab", shortcut_key: "Ctrl+T", description: "Opens a new browser tab"},
        {category: "Navigation", command_name: "New Window", shortcut_key: "Ctrl+N", description: "Opens a new browser window"},
        {category: "Navigation", command_name: "Close Tab", shortcut_key: "Ctrl+W", description: "Closes the current tab"},
        {category: "Navigation", command_name: "Reopen Closed Tab", shortcut_key: "Ctrl+Shift+T", description: "Reopens the last closed tab"}
    ]
    shortcuts_db["chrome.exe"] := chrome_shortcuts
    
    ; Notepad text editor shortcuts
    notepad_shortcuts := [
        {category: "File", command_name: "New", shortcut_key: "Ctrl+N", description: "Creates a new document"},
        {category: "File", command_name: "Open", shortcut_key: "Ctrl+O", description: "Opens an existing document"},
        {category: "File", command_name: "Save", shortcut_key: "Ctrl+S", description: "Saves the current document"},
        {category: "Edit", command_name: "Find", shortcut_key: "Ctrl+F", description: "Opens the find dialog"}
    ]
    shortcuts_db["notepad.exe"] := notepad_shortcuts
    
    ; Microsoft Word shortcuts
    word_shortcuts := [
        {category: "Text Formatting", command_name: "Bold", shortcut_key: "Ctrl+B", description: "Makes selected text bold"},
        {category: "Text Formatting", command_name: "Italic", shortcut_key: "Ctrl+I", description: "Makes selected text italic"},
        {category: "Text Formatting", command_name: "Underline", shortcut_key: "Ctrl+U", description: "Underlines selected text"},
        {category: "File Operations", command_name: "Save", shortcut_key: "Ctrl+S", description: "Saves the current document"}
    ]
    shortcuts_db["winword.exe"] := word_shortcuts
}

/**
 * Create example accessibility compatibility data
 * @return {Void}
 */
AddExampleCompatibility() {
    global at_compat_db
    
    ; Chrome accessibility compatibility
    chrome_compat := [
        {
            at_software: "NVDA Screen Reader", 
            compatibility_level: "Good",
            known_issues: "Some dynamic content may not be announced automatically",
            workarounds: "Use NVDA's browse mode and element navigation"
        },
        {
            at_software: "JAWS Screen Reader", 
            compatibility_level: "Good",
            known_issues: "Custom controls sometimes lack proper labels",
            workarounds: "Use JAWS cursor for unlabeled elements"
        }
    ]
    at_compat_db["chrome.exe"] := chrome_compat
    
    ; Microsoft Word accessibility compatibility
    word_compat := [
        {
            at_software: "NVDA Screen Reader", 
            compatibility_level: "Excellent",
            known_issues: "Complex layout may be difficult to navigate",
            workarounds: "Use document navigation shortcuts (Ctrl+Arrow keys)"
        }
    ]
    at_compat_db["winword.exe"] := word_compat
}

/**
 * Check for multiple script instances that might cause database locking
 * @return {Void}
 */
CheckForMultipleInstances() {
    instanceCount := 0
    currentPID := DllCall("GetCurrentProcessId")
    
    ; Query WMI for AutoHotkey processes
    for process in ComObjGet("winmgmts:").ExecQuery("SELECT * FROM Win32_Process WHERE Name='AutoHotkey64.exe' OR Name='AutoHotkey32.exe' OR Name LIKE '%" A_ScriptName "%'") {
        if (process.ProcessId != currentPID) {
            instanceCount++
        }
    }
    
    if (instanceCount > 0) {
        MsgBox("Warning: " (instanceCount + 1) " instances of this script detected.`n`nMultiple instances can cause database locking issues.", "Multiple Instances", 48)
    }
}

/**
 * Resolve database locking issues through various recovery methods
 * @return {Boolean} True if database is now accessible
 */
ResolveDatabaseLocking() {
    global DB_FILE
    
    if (!FileExist(DB_FILE)) {
        return true
    }
    
    ; Create backup before attempting recovery
    backupFile := DB_FILE . ".backup." . FormatTime(, "yyyyMMdd_HHmmss")
    try {
        FileCopy(DB_FILE, backupFile)
    } catch {
        return false
    }
    
    ; Remove SQLite temporary files that might be causing locks
    try {
        if FileExist(DB_FILE . "-journal")
            FileDelete(DB_FILE . "-journal")
        if FileExist(DB_FILE . "-wal")
            FileDelete(DB_FILE . "-wal")
        if FileExist(DB_FILE . "-shm")
            FileDelete(DB_FILE . "-shm")
    } catch {
        ; Temporary files are locked by another process
    }
    
    ; Allow time for file handles to close
    Sleep(2000)
    
    ; Test database accessibility
    try {
        db := SQLite_Open(DB_FILE)
        if (db) {
            SQLite_Close(db)
            return true
        }
    } catch {
        ; Database is still locked
    }
    
    return false
}

/**
 * Initialize database system with enhanced error handling and embedded file support
 * @return {Void}
 */
InitDatabase() {
    global shortcuts_db, DB_FILE, INSTALL_DIR, debugMode
    
    ; Debug: Show current paths
    if (debugMode) {
        MsgBox("Debug Info:`n`nINSTALL_DIR: " INSTALL_DIR "`nScript Dir: " A_ScriptDir "`nWorking Dir: " A_WorkingDir, "Database Debug", 64)
    }
    
    ; Ensure INSTALL_DIR is set
    if (!INSTALL_DIR || INSTALL_DIR = "") {
        INSTALL_DIR := GetInstallationPath()
        if (debugMode) {
            MsgBox("INSTALL_DIR was empty, set to: " INSTALL_DIR, "Database Debug", 64)
        }
    }
    
    ; Set database file path
    DB_FILE := INSTALL_DIR "\appinfo.db"
    
    ; Check if installation directory exists, create if needed
    if (!DirExist(INSTALL_DIR)) {
        try {
            DirCreate(INSTALL_DIR)
            ; Create subdirectories
            if (!DirExist(INSTALL_DIR "\Backups"))
                DirCreate(INSTALL_DIR "\Backups")
            if (!DirExist(INSTALL_DIR "\Exports"))
                DirCreate(INSTALL_DIR "\Exports")
            if (!DirExist(INSTALL_DIR "\Logs"))
                DirCreate(INSTALL_DIR "\Logs")
                
            if (debugMode) {
                MsgBox("Created installation directory: " INSTALL_DIR, "Database Debug", 64)
            }
        } catch Error as e {
            MsgBox("Error creating installation directory: " e.Message "`n`nUsing script directory instead.", "Database Error", 48)
            INSTALL_DIR := A_ScriptDir
            DB_FILE := INSTALL_DIR "\appinfo.db"
        }
    }
    
    ; Copy embedded files from temp directory to installation directory if needed
    ; This handles the case where FileInstall extracted files to temp directory
    if (A_IsCompiled) {
        tempDir := A_Temp "\Keyfinder_Temp"
        
        ; Copy sqlite3.dll if it doesn't exist in INSTALL_DIR
        dllSource := tempDir "\sqlite3.dll"
        dllDest := INSTALL_DIR "\sqlite3.dll"
        dllCopied := false
        if (FileExist(dllSource) && !FileExist(dllDest)) {
            try {
                FileCopy(dllSource, dllDest, 1)
                dllCopied := true
                if (debugMode) {
                    MsgBox("Copied sqlite3.dll from temp directory to installation directory", "Database Debug", 64)
                }
            } catch Error as e {
                if (debugMode) {
                    MsgBox("Could not copy sqlite3.dll: " e.Message, "Database Debug", 48)
                }
            }
        }
        
        ; Copy appinfo.db if it doesn't exist in INSTALL_DIR
        dbSource := tempDir "\appinfo.db"
        dbCopied := false
        if (FileExist(dbSource) && !FileExist(DB_FILE)) {
            try {
                FileCopy(dbSource, DB_FILE, 0)  ; Don't overwrite existing database
                dbCopied := true
                if (debugMode) {
                    MsgBox("Copied appinfo.db from temp directory to installation directory", "Database Debug", 64)
                }
            } catch Error as e {
                if (debugMode) {
                    MsgBox("Could not copy appinfo.db: " e.Message, "Database Debug", 48)
                }
            }
        }
        
        ; CLEANUP: Delete extracted files from temp directory
        ; These are in the hidden temp folder anyway, but good to clean up
        try {
            if (dllCopied || (FileExist(dllDest) && FileExist(dllSource))) {
                FileDelete(dllSource)
                if (debugMode) {
                    MsgBox("Cleaned up sqlite3.dll from temp: " dllSource, "Cleanup", 64)
                }
            }
            
            if (dbCopied || (FileExist(DB_FILE) && FileExist(dbSource))) {
                FileDelete(dbSource)
                if (debugMode) {
                    MsgBox("Cleaned up appinfo.db from temp: " dbSource, "Cleanup", 64)
                }
            }
            
            ; Try to remove temp directory if empty
            if (DirExist(tempDir)) {
                try {
                    DirDelete(tempDir)
                } catch {
                    ; Ignore errors - temp directory will be cleaned by Windows eventually
                }
            }
        } catch Error as e {
            ; Non-critical - temp files are hidden from user anyway
            if (debugMode) {
                MsgBox("Could not clean up temp files: " e.Message, "Cleanup Warning", 48)
            }
        }
    }
    
    ; Debug: Show final paths
    if (debugMode) {
        MsgBox("Final paths:`n`nINSTALL_DIR: " INSTALL_DIR "`nDB_FILE: " DB_FILE "`nDB exists: " (FileExist(DB_FILE) ? "Yes" : "No"), "Database Debug", 64)
    }
    
    ; Check for multiple instances that could cause database conflicts
    CheckForMultipleInstances()
    
    isSqliteAvailable := InitializeSQLite()
    
    if (debugMode) {
        MsgBox("SQLite Available: " (isSqliteAvailable ? "Yes" : "No") "`nSQLite DLL exists: " (FileExist(A_ScriptDir "\sqlite3.dll") ? "Yes" : "No"), "Database Debug", 64)
    }
    
    if (isSqliteAvailable) {
        dbExists := FileExist(DB_FILE)
        
        ; Always try to create database if it doesn't exist
        if (!dbExists) {
            try {
                CreateEmptyDatabase()
                PopulateInitialData()
                if (debugMode) {
                    MsgBox("Created new database and populated with initial data", "Database Debug", 64)
                }
            } catch Error as e {
                MsgBox("Could not create initial database: " e.Message "`nUsing backup file instead.", "Database Warning", 48)
                LoadShortcutsFromFile()
            }
        }
        
        ; Attempt to load from SQLite database
        try {
            db := SQLite_Open(DB_FILE)
            
            if (db) {
                ; Ensure tables exist with proper schema
                SQLite_Exec(db, "CREATE TABLE IF NOT EXISTS shortcuts (id INTEGER PRIMARY KEY, process_name TEXT, command_name TEXT, shortcut_key TEXT, category TEXT, description TEXT);")
                SQLite_Exec(db, "CREATE TABLE IF NOT EXISTS at_compatibility (id INTEGER PRIMARY KEY, process_name TEXT, at_software TEXT, compatibility_level TEXT, known_issues TEXT, workarounds TEXT);")
                SQLite_Exec(db, "CREATE TABLE IF NOT EXISTS at_usage (id INTEGER PRIMARY KEY, process_name TEXT, window_class TEXT, at_software TEXT, timestamp TEXT);")
                
                ; Create performance indexes
                SQLite_Exec(db, "CREATE INDEX IF NOT EXISTS idx_shortcuts_process ON shortcuts(process_name);")
                SQLite_Exec(db, "CREATE INDEX IF NOT EXISTS idx_at_compat_process ON at_compatibility(process_name);")
                
                LoadDataFromDatabase(db)
                SQLite_Close(db)
                
                if (debugMode) {
                    MsgBox("Successfully loaded data from SQLite database", "Database Debug", 64)
                }
            } else {
                if (debugMode) {
                    MsgBox("Failed to open SQLite database, using backup file", "Database Debug", 48)
                }
                LoadShortcutsFromFile()
            }
        } catch Error as e {
            MsgBox("Database error: " e.Message "`nUsing backup file instead.", "Database Error", 48)
            LoadShortcutsFromFile()
        }
    } else {
        if (debugMode) {
            MsgBox("SQLite not available, using backup file system", "Database Debug", 48)
        }
        LoadShortcutsFromFile()
    }
    
    ; Ensure we have some data to work with
    if (shortcuts_db.Count = 0) {
        AddExampleShortcuts()
        if (debugMode) {
            MsgBox("Added example shortcuts to memory database", "Database Debug", 64)
        }
    }
    
    if (at_compat_db.Count = 0) {
        AddExampleCompatibility()
    }
    
    ; Final debug info
    if (debugMode) {
        MsgBox("Database initialization complete:`n`nShortcuts in memory: " shortcuts_db.Count " apps`nCompatibility data: " at_compat_db.Count " apps", "Database Debug", 64)
    }
}

/**
 * Save in-memory database to SQLite with transaction support
 * @param {Integer} db - Database handle (optional, will open if not provided)
 * @return {Boolean} True if save operation successful
 */
SaveMemoryToDatabase(db := 0) {
    global shortcuts_db, DB_FILE
    
    closeDbWhenDone := false
    if (!db) {
        if (!InitializeSQLite())
            return false
            
        db := SQLite_Open(DB_FILE)
        if (!db)
            return false
            
        closeDbWhenDone := true
    }
    
    try {
        ; Use transaction for atomic operation
        SQLite_Exec(db, "BEGIN TRANSACTION;")
        SQLite_Exec(db, "DELETE FROM shortcuts;")
        
        totalSaved := 0
        
        ; Save all shortcuts from memory to database
        for processName, shortcuts in shortcuts_db {
            for _, shortcut in shortcuts {
                AddShortcutToDatabase(db, 
                    processName, 
                    shortcut.HasOwnProp("command_name") ? shortcut.command_name : "", 
                    shortcut.HasOwnProp("shortcut_key") ? shortcut.shortcut_key : "", 
                    shortcut.HasOwnProp("category") ? shortcut.category : "General", 
                    shortcut.HasOwnProp("description") ? shortcut.description : ""
                )
                totalSaved++
            }
        }
        
        SQLite_Exec(db, "COMMIT;")
        return true
    } catch Error as e {
        try SQLite_Exec(db, "ROLLBACK;")
        MsgBox("Error saving to database: " e.Message, "Database Error", 16)
        return false
    } finally {
        if (closeDbWhenDone && db)
            SQLite_Close(db)
    }
}

/**
 * Save shortcuts to backup INI file for fallback storage
 * @return {Boolean} True if backup file saved successfully
 */
SaveShortcutsToFile() {
    global shortcuts_db, INSTALL_DIR
    
    fileContent := ""
    
    ; Generate INI-style backup file content
    for processName, shortcuts in shortcuts_db {
        fileContent .= "[" processName "]`n"
        
        for i, shortcut in shortcuts {
            fileContent .= "shortcut" i "_category=" (shortcut.HasOwnProp("category") ? shortcut.category : "General") "`n"
            fileContent .= "shortcut" i "_command=" (shortcut.HasOwnProp("command_name") ? shortcut.command_name : "") "`n"
            fileContent .= "shortcut" i "_keys=" (shortcut.HasOwnProp("shortcut_key") ? shortcut.shortcut_key : "") "`n"
            fileContent .= "shortcut" i "_desc=" (shortcut.HasOwnProp("description") ? shortcut.description : "") "`n"
        }
        fileContent .= "`n"
    }
    
    backupFile := INSTALL_DIR "\Backups\shortcuts_backup.ini"
    try {
        if FileExist(backupFile)
            FileDelete(backupFile)
        
        FileAppend(fileContent, backupFile)
        return true
    } catch Error as e {
        MsgBox("Error saving shortcuts to file: " e.Message, "Error", 16)
        return false
    }
}

/**
 * Add shortcut to both memory and persistent storage
 * @param {String} processName - Process executable name
 * @param {String} windowClass - Window class name (for future use)
 * @param {String} commandName - Human-readable command name
 * @param {String} shortcutKey - Keyboard shortcut combination
 * @param {String} description - Detailed description of the shortcut
 * @param {String} category - Shortcut category (default "General")
 * @return {Boolean} True if shortcut added successfully
 */
AddShortcutToDB(processName, windowClass, commandName, shortcutKey, description, category := "General") {
    global shortcuts_db, DB_FILE, lastDetectedHwnd
    
    ; Add to in-memory database
    if !shortcuts_db.Has(processName)
        shortcuts_db[processName] := []
        
    shortcuts_db[processName].Push({
        category: category,
        command_name: commandName,
        shortcut_key: shortcutKey,
        description: description
    })
    
    ; Save to backup file immediately
    SaveShortcutsToFile()
    
    ; Attempt to save to SQLite database
    if (InitializeSQLite()) {
        try {
            db := SQLite_Open(DB_FILE)
            if db {
                success := AddShortcutToDatabase(db, processName, commandName, shortcutKey, category, description)
                SQLite_Close(db)
                
                if !success {
                    MsgBox("Failed to add shortcut to SQLite but saved to backup file", "Database Warning", 48)
                }
            } else {
                MsgBox("Could not open SQLite database but saved to backup file", "Database Warning", 48)
            }
        } catch Error as e {
            MsgBox("Database error: " e.Message "`nShortcuts saved to backup file instead.", "Database Warning", 48)
        }
    }
    
    return true
}

/**
 * Display a dialog to capture a new hotkey combination.
 * @param {Object} targetEdit - The Edit control to update with the new hotkey.
 * @param {String} hotkeyName - The name of the hotkey being remapped ('capture' or 'search').
 * @return {Void}
 */
CaptureKeyForRemap(targetEdit, hotkeyName) {
    CaptureGui := Gui("+Owner" . targetEdit.Gui.Hwnd, "Set New Hotkey")
    ApplyDarkModeToGUI(CaptureGui, "Set Hotkey")
    colors := GetThemeColors()
    isDark := IsWindowsDarkMode()

    instructionText := CaptureGui.Add("Text", "x10 y10 w280 center", "Press the new key combination for " hotkeyName " and then click 'Set'.")
    ApplyControlTheming(instructionText, "Text", colors, isDark)
    
    CapturedKeyEdit := CaptureGui.Add("Edit", "x50 y40 w200 ReadOnly", "")
    ApplyControlTheming(CapturedKeyEdit, "Edit", colors, isDark)
    
    SetButton := CaptureGui.Add("Button", "Default x50 y80 w100", "Set")
    ApplyControlTheming(SetButton, "Button", colors, isDark)
    
    CancelButton := CaptureGui.Add("Button", "x160 y80 w100", "Cancel")
    ApplyControlTheming(CancelButton, "Button", colors, isDark)

    pressedKeys := Map()

    UpdateKeystrokeDisplay(keysMap) {
        modifiers := []
        regularKeys := []
        
        for key, _ in keysMap {
            if (key = "LCtrl" || key = "RCtrl")
                AddToArrayIfMissing(modifiers, "Ctrl")
            else if (key = "LAlt" || key = "RAlt")
                AddToArrayIfMissing(modifiers, "Alt")
            else if (key = "LShift" || key = "RShift")
                AddToArrayIfMissing(modifiers, "Shift")
            else if (key = "LWin" || key = "RWin")
                AddToArrayIfMissing(modifiers, "Win")
            else
                AddToArrayIfMissing(regularKeys, key)
        }
        
        keystrokeText := ""
        if (modifiers.Length > 0)
            keystrokeText := Join(modifiers, "+")
            
        if (regularKeys.Length > 0) {
            if (keystrokeText != "")
                keystrokeText .= "+"
            keystrokeText .= Join(regularKeys, "+")
        }
        
        CapturedKeyEdit.Value := keystrokeText
        ApplyControlTheming(CapturedKeyEdit, "Edit", colors, isDark)
    }

    KeyDown(key, *) {
        pressedKeys[key] := true
        UpdateKeystrokeDisplay(pressedKeys)
    }

    KeyUp(key, *) {
        if pressedKeys.Has(key)
            pressedKeys.Delete(key)
    }

    RegisterHotkeysForCapture() {
        HotIfWinActive("ahk_id " . CaptureGui.Hwnd)
        keysToRegister := ["LCtrl", "RCtrl", "LAlt", "RAlt", "LShift", "RShift", "LWin", "RWin", "Tab", "Space", "Delete", "Home", "End", "PgUp", "PgDn", "Insert", "NumpadAdd", "NumpadSub", "NumpadMult", "NumpadDiv", "Left", "Right", "Up", "Down"]
        alphabet := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Loop Parse, alphabet
            keysToRegister.Push(A_LoopField)
        Loop 10
            keysToRegister.Push(A_Index-1)
        Loop 12
            keysToRegister.Push("F" . A_Index)
        
        for key in keysToRegister {
            Hotkey("~*" . key, KeyDown.Bind(key))
            Hotkey("~*" . key . " up", KeyUp.Bind(key))
        }
    }

    SetNewHotkey(*) {
        if (CapturedKeyEdit.Value != "") {
            targetEdit.Value := CapturedKeyEdit.Value
            CaptureGui.Destroy()
        } else {
            MsgBox("No hotkey captured. Please press a key combination.", "Error", 48)
        }
    }

    CancelCapture(*) {
        CaptureGui.Destroy()
    }

    SetButton.OnEvent("Click", SetNewHotkey)
    CancelButton.OnEvent("Click", CancelCapture)
    CaptureGui.OnEvent("Close", CancelCapture)
    
    RegisterHotkeysForCapture()
    CaptureGui.Show("w300 h120")
}
;==============================================================================
; PREFERENCES MANAGEMENT SYSTEM WITH ENHANCED THEME CHECKING
;==============================================================================

/**
 * Display application preferences configuration dialog with enhanced theme checking
 * @return {Void}
 */
ShowPreferencesDialog(*) {
    global SETTINGS_FILE, SCRIPT_NAME, MyGui, captureHotkey, searchHotkey

    ; Ensure the theme is up-to-date before showing the dialog
    ForceThemeCheck()    
    
    try {
        PrefsGui := CreateManagedDialog("+Owner" MyGui.Hwnd " +Resize +MinSize370x550", SCRIPT_NAME " Settings")
    } catch {
        PrefsGui := Gui("+Owner" MyGui.Hwnd " +MinSize370x550", SCRIPT_NAME " Settings")
        ApplyDarkModeToGUI(PrefsGui, "Settings Dialog")
    }
    PrefsGui.SetFont("s10")
    colors := GetThemeColors()
    isDark := IsWindowsDarkMode()
    
    HeaderText := PrefsGui.Add("Text", "x10 y10 w350 center", "Application Settings")
    ApplyControlTheming(HeaderText, "Text", colors, isDark)
    
    ; Load current settings from configuration file
    try {
        currentStartup := (IniRead(SETTINGS_FILE, "General", "StartupEnabled", "false") = "true")
        currentDesktop := (IniRead(SETTINGS_FILE, "General", "DesktopShortcut", "false") = "true")
        currentQuiet := (IniRead(SETTINGS_FILE, "General", "QuietMode", "false") = "true")
        currentDebug := (IniRead(SETTINGS_FILE, "General", "DebugMode", "false") = "true")
    } catch {
        currentStartup := false
        currentDesktop := false
        currentQuiet := false
        currentDebug := false
    }
    
    ; System Integration section
    systemGroupBox := PrefsGui.Add("GroupBox", "x10 y40 w350 h140", "System Integration")
    ApplyControlTheming(systemGroupBox, "GroupBox", colors, isDark)
    
    startupCheck := PrefsGui.Add("Checkbox", "x20 y65 w320", "&Start with Windows")
    startupCheck.Value := currentStartup
    ApplyControlTheming(startupCheck, "Button", colors, isDark)
    
    startupDesc := PrefsGui.Add("Text", "x40 y85 w300", "Automatically start the application when Windows starts")
    ApplyControlTheming(startupDesc, "Text", colors, isDark)
    
    desktopCheck := PrefsGui.Add("Checkbox", "x20 y115 w320", "&Desktop shortcut")
    desktopCheck.Value := currentDesktop
    ApplyControlTheming(desktopCheck, "Button", colors, isDark)
    
    desktopDesc := PrefsGui.Add("Text", "x40 y135 w300", "Show shortcut icon on desktop for easy access")
    ApplyControlTheming(desktopDesc, "Text", colors, isDark)
    
    ; Application Behavior section
    behaviorGroupBox := PrefsGui.Add("GroupBox", "x10 y190 w350 h110", "Application Behavior")
    ApplyControlTheming(behaviorGroupBox, "GroupBox", colors, isDark)
    
    quietCheck := PrefsGui.Add("Checkbox", "x20 y215 w320", "&Quiet mode")
    quietCheck.Value := currentQuiet
    ApplyControlTheming(quietCheck, "Button", colors, isDark)
    
    quietDesc := PrefsGui.Add("Text", "x40 y235 w300", "Reduce notification messages")
    ApplyControlTheming(quietDesc, "Text", colors, isDark)
    
    debugCheck := PrefsGui.Add("Checkbox", "x20 y260 w320", "Debug &mode")
    debugCheck.Value := currentDebug
    ApplyControlTheming(debugCheck, "Button", colors, isDark)
    
    debugDesc := PrefsGui.Add("Text", "x40 y280 w300", "Show additional debug information (for troubleshooting)")
    ApplyControlTheming(debugDesc, "Text", colors, isDark)

    ; Hotkey Remapping section
    hotkeyGroupBox := PrefsGui.Add("GroupBox", "x10 y310 w350 h110", "Hotkey Remapping")
    ApplyControlTheming(hotkeyGroupBox, "GroupBox", colors, isDark)

    captureHotkeyLabel := PrefsGui.Add("Text", "x20 y335 w150", "Capture Window Hotkey:")
    ApplyControlTheming(captureHotkeyLabel, "Text", colors, isDark)
    captureHotkeyEdit := PrefsGui.Add("Edit", "x180 y335 w100 ReadOnly", captureHotkey)
    ApplyControlTheming(captureHotkeyEdit, "Edit", colors, isDark)
    captureHotkeyButton := PrefsGui.Add("Button", "x290 y335 w70", "Set")
    ApplyControlTheming(captureHotkeyButton, "Button", colors, isDark)

    searchHotkeyLabel := PrefsGui.Add("Text", "x20 y375 w150", "Search Hotkey:")
    ApplyControlTheming(searchHotkeyLabel, "Text", colors, isDark)
    searchHotkeyEdit := PrefsGui.Add("Edit", "x180 y375 w100 ReadOnly", searchHotkey)
    ApplyControlTheming(searchHotkeyEdit, "Edit", colors, isDark)
    searchHotkeyButton := PrefsGui.Add("Button", "x290 y375 w70", "Set")
    ApplyControlTheming(searchHotkeyButton, "Button", colors, isDark)

    captureHotkeyButton.OnEvent("Click", (*) => CaptureKeyForRemap(captureHotkeyEdit, "capture"))
    searchHotkeyButton.OnEvent("Click", (*) => CaptureKeyForRemap(searchHotkeyEdit, "search"))
    
    ; Application Management section
    managementGroupBox := PrefsGui.Add("GroupBox", "x10 y430 w350 h60", "Application Management")
    ApplyControlTheming(managementGroupBox, "GroupBox", colors, isDark)
    
    uninstallButton := PrefsGui.Add("Button", "x20 y455 w150", "&Uninstall Application")
    ApplyControlTheming(uninstallButton, "Button", colors, isDark)
    
    uninstallDesc := PrefsGui.Add("Text", "x180 y458 w170", "Remove " SCRIPT_NAME " from system")
    ApplyControlTheming(uninstallDesc, "Text", colors, isDark)
    
    ; Action buttons
    savePrefsButton := PrefsGui.Add("Button", "Default x70 y510 w100", "&Save")
    ApplyControlTheming(savePrefsButton, "Button", colors, isDark)
    
    cancelPrefsButton := PrefsGui.Add("Button", "x180 y510 w100", "&Cancel")
    ApplyControlTheming(cancelPrefsButton, "Button", colors, isDark)

    ; Store control references in the GUI object for resize handler
    PrefsGui.controls := {
        systemGroupBox: systemGroupBox,
        startupCheck: startupCheck,
        startupDesc: startupDesc,
        desktopCheck: desktopCheck,
        desktopDesc: desktopDesc,
        behaviorGroupBox: behaviorGroupBox,
        quietCheck: quietCheck,
        quietDesc: quietDesc,
        debugCheck: debugCheck,
        debugDesc: debugDesc,
        hotkeyGroupBox: hotkeyGroupBox,
        captureHotkeyLabel: captureHotkeyLabel,
        captureHotkeyEdit: captureHotkeyEdit,
        captureHotkeyButton: captureHotkeyButton,
        searchHotkeyLabel: searchHotkeyLabel,
        searchHotkeyEdit: searchHotkeyEdit,
        searchHotkeyButton: searchHotkeyButton,
        managementGroupBox: managementGroupBox,
        uninstallButton: uninstallButton,
        uninstallDesc: uninstallDesc,
        savePrefsButton: savePrefsButton,
        cancelPrefsButton: cancelPrefsButton
    }

    SavePreferences(*) {
        try {
            newStartup := startupCheck.Value
            newDesktop := desktopCheck.Value
            newQuiet := quietCheck.Value
            newDebug := debugCheck.Value
            
            ; Apply system integration changes
            if (newStartup != currentStartup) {
                SetStartupIntegration(newStartup)
            }
            
            if (newDesktop != currentDesktop) {
                SetDesktopShortcut(newDesktop)
            }
            
            ; Save behavior settings
            IniWrite(newQuiet ? "true" : "false", SETTINGS_FILE, "General", "QuietMode")
            IniWrite(newDebug ? "true" : "false", SETTINGS_FILE, "General", "DebugMode")
            
            ; Save Hotkey settings
            IniWrite(captureHotkeyEdit.Value, SETTINGS_FILE, "Hotkeys", "CaptureWindow")
            IniWrite(searchHotkeyEdit.Value, SETTINGS_FILE, "Hotkeys", "SearchShortcut")

            ; Update global variables
            global quietMode, debugMode, captureHotkey, searchHotkey
            
            quietMode := newQuiet
            debugMode := newDebug
            captureHotkey := captureHotkeyEdit.Value
            searchHotkey := searchHotkeyEdit.Value

            RegisterGlobalHotkeys()
            
            MsgBox("Settings saved successfully. New hotkeys are now active.", "Settings", 64)
            PrefsGui.Destroy()
            
        } catch Error as e {
            MsgBox("Error saving settings: " e.Message, "Settings Error", 16)
        }
    }
    
    ; Resize handler for preferences dialog
    ResizePreferencesDialog(GuiObj, MinMax, Width, Height) {
        ; Enforce minimum window size
        minWidth := 370
        minHeight := 550
        
        if (Width < minWidth)
            Width := minWidth
        if (Height < minHeight)
            Height := minHeight
            
        ; Calculate content dimensions
        margin := 10
        groupWidth := Width - 2 * margin
        labelWidth := groupWidth - 40

        try {
            controls := GuiObj.controls
            
            ; Resize System Integration section
            controls.systemGroupBox.Move(margin, 40, groupWidth, 140)
            controls.startupCheck.Move(margin + 10, 65, groupWidth - 20)
            controls.startupDesc.Move(margin + 30, 85, labelWidth)
            controls.desktopCheck.Move(margin + 10, 115, groupWidth - 20)
            controls.desktopDesc.Move(margin + 30, 135, labelWidth)

            ; Resize Application Behavior section
            controls.behaviorGroupBox.Move(margin, 190, groupWidth, 110)
            controls.quietCheck.Move(margin + 10, 215, groupWidth - 20)
            controls.quietDesc.Move(margin + 30, 235, labelWidth)
            controls.debugCheck.Move(margin + 10, 260, groupWidth - 20)
            controls.debugDesc.Move(margin + 30, 280, labelWidth)

            ; Resize Hotkey section
            controls.hotkeyGroupBox.Move(margin, 310, groupWidth, 110)
            controls.captureHotkeyLabel.Move(margin + 10, 335, 150)
            controls.captureHotkeyEdit.Move(margin + 170, 335, 100)
            controls.captureHotkeyButton.Move(Width - 80, 335, 70)
            controls.searchHotkeyLabel.Move(margin + 10, 375, 150)
            controls.searchHotkeyEdit.Move(margin + 170, 375, 100)
            controls.searchHotkeyButton.Move(Width - 80, 375, 70)
            
            ; Resize Management section
            controls.managementGroupBox.Move(margin, 430, groupWidth, 60)
            controls.uninstallButton.Move(margin + 10, 455, 150)
            controls.uninstallDesc.Move(margin + 170, 458, groupWidth - 180)
            
            ; Resize action buttons
            buttonY := Height - 40
            controls.savePrefsButton.Move((Width / 2) - 110, buttonY, 100)
            controls.cancelPrefsButton.Move((Width / 2) + 10, buttonY, 100)
            
            ; Force redraw
            GuiObj.Redraw()
            
        } catch Error as e {
            ; Silently handle resize errors
        }
    }
    
    ; Set the resize event handler
    PrefsGui.OnEvent("Size", ResizePreferencesDialog)
    
    ; Event handlers
    savePrefsButton.OnEvent("Click", SavePreferences)
    cancelPrefsButton.OnEvent("Click", (*) => PrefsGui.Destroy())
    uninstallButton.OnEvent("Click", StartUninstall)
    PrefsGui.OnEvent("Close", (*) => PrefsGui.Destroy())
    
    RegisterGuiForEscapeHandling(PrefsGui, (*) => PrefsGui.Destroy())
    
    PrefsGui.Show("w370 h550")
    ForceWindowToForeground(PrefsGui.Hwnd)
    
    ; Force theme refresh after showing
    Sleep(100)
    RefreshGUITheme(PrefsGui, "Settings Dialog")
    
    ; Apply enhanced text control theming
    ForceTextControlsRetheming(PrefsGui)
}

;==============================================================================
; GUI SYSTEM - MAIN INTERFACE WITH ENHANCED RESPONSIVE DARK MODE
;==============================================================================

/**
 * Enhanced RefreshMainGUITheme with extra Text control focus
 */
RefreshMainGUITheme() {
    global MyGui, ShortcutsText, CategoryFilter
    global AddAppButton, PreferencesButton, CloseButton, HelpButton, ExportButton
    global FilterLabel, ShortcutsHeader
    
    try {
        if (!MyGui || !MyGui.Hwnd || !WinExist("ahk_id " MyGui.Hwnd)) {
            return
        }
        
        ; Apply theme to main window
        ApplyDarkModeToGUI(MyGui, "Main Application Window")
        colors := GetThemeColors()
        isDark := IsWindowsDarkMode()
        
        ; ENHANCED: Focus specifically on Text controls first
        if (IsObject(ShortcutsHeader)) {
            ApplyControlTheming(ShortcutsHeader, "Text", colors, isDark)
            ; Force immediate refresh
            ForceTextControlRefresh(ShortcutsHeader)
        }
        if (IsObject(FilterLabel)) {
            ApplyControlTheming(FilterLabel, "Text", colors, isDark)
            ; Force immediate refresh
            ForceTextControlRefresh(FilterLabel)
        }
        
        ; Refresh other main window controls
        if (IsObject(ShortcutsText))
            ApplyControlTheming(ShortcutsText, "Edit", colors, isDark)
        if (IsObject(CategoryFilter))
            ApplyControlTheming(CategoryFilter, "DropDownList", colors, isDark)
        if (IsObject(AddAppButton))
            ApplyControlTheming(AddAppButton, "Button", colors, isDark)
        if (IsObject(ExportButton))
            ApplyControlTheming(ExportButton, "Button", colors, isDark)
        if (IsObject(PreferencesButton))
            ApplyControlTheming(PreferencesButton, "Button", colors, isDark)
        if (IsObject(CloseButton))
            ApplyControlTheming(CloseButton, "Button", colors, isDark)
        if (IsObject(HelpButton))
            ApplyControlTheming(HelpButton, "Button", colors, isDark)
        
        ; Force re-theming of all text controls
        ForceTextControlsRetheming(MyGui)

        ; Force a complete and immediate repaint
        try {
            DllCall("RedrawWindow", "Ptr", MyGui.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x0001 | 0x0004 | 0x0010)
        } catch {
            MyGui.Redraw()
        }
        
        ; Schedule an additional refresh for stubborn Text controls
        SetTimer(AdditionalTextControlRefresh, -300)

    } catch Error as e {
        ; Silently handle refresh errors
    }
}

/**
 * Additional refresh specifically for Text controls that might be stubborn
 */
AdditionalTextControlRefresh() {
    global ShortcutsHeader, FilterLabel
    
    try {
        colors := GetThemeColors()
        isDark := IsWindowsDarkMode()
        
        if (IsObject(ShortcutsHeader)) {
            ; Re-apply theming
            ApplyControlTheming(ShortcutsHeader, "Text", colors, isDark)
            ; Force refresh
            ForceTextControlRefresh(ShortcutsHeader)
        }
        
        if (IsObject(FilterLabel)) {
            ; Re-apply theming
            ApplyControlTheming(FilterLabel, "Text", colors, isDark)
            ; Force refresh
            ForceTextControlRefresh(FilterLabel)
        }
        
    } catch {
        ; Additional refresh failed
    }
}

/**
 * Initialize and create the main application GUI interface with enhanced responsive dark mode support
 * @return {Void}
 */
InitializeGUIWithThemeSupport() {
    global MyGui, ShortcutsText, CategoryFilter
    global AddAppButton, PreferencesButton, CloseButton, HelpButton, ExportButton
    global FilterLabel, ShortcutsHeader
    
    ; Create main GUI window with resize capabilities
    MyGui := Gui("+Resize +MinSize520x450", "Application Shortcuts")
    MyGui.SetFont("s10")
    
    ; Apply dark mode theming to main window
    ApplyDarkModeToGUI(MyGui, "Main Application Window")
    colors := GetThemeColors()
    isDark := IsWindowsDarkMode()
    
    ; Store reference to refresh function in GUI object
    MyGui.RefreshTheme := RefreshMainGUITheme
    
    ; Header section
    ShortcutsHeader := MyGui.Add("Text", "x10 y10 w500 center", "Application Shortcuts")
    ShortcutsHeader.SetFont("bold s11")
    ApplyControlTheming(ShortcutsHeader, "Text", colors, isDark)
    
    ; Main content area - shortcuts display
    ShortcutsText := MyGui.Add("Edit", "x10 y35 w500 h300 ReadOnly VScroll", "Capture an application to view available shortcuts.")
    ApplyControlTheming(ShortcutsText, "Edit", colors, isDark)
    
    ; Filter controls section
    FilterLabel := MyGui.Add("Text", "x15 y345 w100", "Filter Category:")
    ApplyControlTheming(FilterLabel, "Text", colors, isDark)
    
    CategoryFilter := MyGui.Add("DropDownList", "x15 y365 w150 Visible", ["All Categories"])
    ApplyControlTheming(CategoryFilter, "DropDownList", colors, isDark)
    CategoryFilter.Choose(1)
    CategoryFilter.OnEvent("Change", UpdateShortcutsHandler)
    
    ; Action buttons - primary functions
    AddAppButton := MyGui.Add("Button", "x180 y345 w120", "Add to Database")
    ApplyControlTheming(AddAppButton, "Button", colors, isDark)
    AddAppButton.OnEvent("Click", AddAppHandler)
    
    ExportButton := MyGui.Add("Button", "x320 y345 w100", "&Export")
    ApplyControlTheming(ExportButton, "Button", colors, isDark)
    ExportButton.OnEvent("Click", ExportHandler)
    
    ; Bottom navigation buttons
    PreferencesButton := MyGui.Add("Button", "x70 y400 w100", "&Settings")
    ApplyControlTheming(PreferencesButton, "Button", colors, isDark)
    
    CloseButton := MyGui.Add("Button", "x180 y400 w100", "&Close")
    ApplyControlTheming(CloseButton, "Button", colors, isDark)
    
    HelpButton := MyGui.Add("Button", "x290 y400 w100", "&Help")
    ApplyControlTheming(HelpButton, "Button", colors, isDark)
    
    ; Set up event handlers
    PreferencesButton.OnEvent("Click", PreferencesHandler)
    CloseButton.OnEvent("Click", CloseHandler)
    HelpButton.OnEvent("Click", HelpHandler)
    MyGui.OnEvent("Close", CloseHandler)
    
    ; Assign resize event handler for responsive layout
    MyGui.OnEvent("Size", OnGUIResize)
    
    ; Register Escape key to hide main window
    RegisterGuiForEscapeHandling(MyGui, (*) => MyGui.Hide())
    
    ; Ensure controls are visible and properly rendered
    CategoryFilter.Visible := true
    CategoryFilter.Redraw()
    
    ; Apply enhanced text control theming
    ForceTextControlsRetheming(MyGui)
    
    ; Start with window hidden
    MyGui.Hide()
}

;==============================================================================
; GUI EVENT HANDLERS WITH ENHANCED THEME CHECKING
;==============================================================================

/**
 * Handle Add Application button click
 * @return {Void}
 */
AddAppHandler(*) {
    if (lastDetectedHwnd = 0 || !WinExist("ahk_id " lastDetectedHwnd)) {
        MsgBox("Please capture an application first using ` & 2.", "No Application Detected", 48)
        return
    }
    ShowAddAppDialog()
}

/**
 * Handle Preferences button click
 * @return {Void}
 */
PreferencesHandler(*) {
    ShowPreferencesDialog()
}

/**
 * Handle Close button click and window close events
 * @return {Void}
 */
CloseHandler(*) {
    global MyGui
    
    ; Hide window instead of closing application entirely
    try {
        MyGui.Hide()
    } catch {
        ; Window might already be closed
    }
}

/**
 * Handle Help button click
 * @return {Void}
 */
HelpHandler(*) {
    ShowHelpDialog()
}

/**
 * Handle Export button click
 * @return {Void}
 */
ExportHandler(*) {
    ShowExportDialog()
}

/**
 * Handle category filter change events
 * @return {Void}
 */
UpdateShortcutsHandler(*) {
    UpdateShortcutsWithFilter()
}

/**
 * Enhanced GUI resize handler for responsive layout management
 * @param {Object} MyGui - Main GUI object being resized
 * @param {Integer} MinMax - Minimize/maximize state indicator
 * @param {Integer} Width - New window width
 * @param {Integer} Height - New window height
 * @return {Void}
 */
OnGUIResize(MyGui, MinMax, Width, Height) {
    global ShortcutsText, CategoryFilter, AddAppButton, PreferencesButton, CloseButton, HelpButton, ExportButton
    global FilterLabel, ShortcutsHeader
    
    ; Enforce minimum window dimensions
    minWidth := 520
    minHeight := 450
    
    if (Width < minWidth)
        Width := minWidth
    if (Height < minHeight)
        Height := minHeight
    
    ; Calculate content area dimensions
    contentWidth := Width - 20
    
    try {
        ; Resize header section
        if (IsObject(ShortcutsHeader)) {
            ShortcutsHeader.Move(10, 10, contentWidth, 25)
        }
        
        ; Calculate main content area height (reserve space for controls)
        shortcutsHeight := Height - 155  ; Total space: 35 (header) + 120 (controls) = 155
        if (shortcutsHeight < 200)  ; Minimum content height
            shortcutsHeight := 200
        
        if (IsObject(ShortcutsText)) {
            ShortcutsText.Move(10, 35, contentWidth, shortcutsHeight)
        }
        
        ; Position filter controls with proper spacing
        filterY := 35 + shortcutsHeight + 10
        
        if (IsObject(FilterLabel)) {
            FilterLabel.Move(15, filterY, 100, 16)
            FilterLabel.Visible := true
        }
        
        if (IsObject(CategoryFilter)) {
            CategoryFilter.Move(15, filterY + 20, 150, 21)
            CategoryFilter.Visible := true
            CategoryFilter.Redraw()
        }
        
        ; Position action buttons
        if (IsObject(AddAppButton)) {
            AddAppButton.Move(180, filterY, 120, 23)
        }
        
        if (IsObject(ExportButton)) {
            ExportButton.Move(contentWidth - 110, filterY, 100, 23)
        }
        
        ; Position bottom navigation buttons (centered)
        buttonY := filterY + 55
        buttonSpacing := 110
        totalButtonWidth := buttonSpacing * 3
        startX := (Width - totalButtonWidth) / 2
        
        if (IsObject(PreferencesButton)) {
            PreferencesButton.Move(startX, buttonY, 100, 25)
        }
        
        if (IsObject(CloseButton)) {
            CloseButton.Move(startX + buttonSpacing, buttonY, 100, 25)
        }
        
        if (IsObject(HelpButton)) {
            HelpButton.Move(startX + (buttonSpacing * 2), buttonY, 100, 25)
        }
        
        ; Force visual refresh to prevent display artifacts
        MyGui.Redraw()
        
    } catch Error as e {
        ; Silently handle control access errors during resize
    }
}

;==============================================================================
; CORE WINDOW CAPTURE FUNCTIONALITY WITH ENHANCED THEME CHECKING
;==============================================================================

/**
 * Enhanced CaptureFocusedWindow with proper theme handling
 * @return {Void}
 */
CaptureFocusedWindow() {
    global lastDetectedHwnd, MyGui
    
    ; Force theme check with stability measures
    ForceThemeCheck()
    
    ; Hide main window temporarily to avoid capturing it
    wasVisible := false
    try {
        if (WinExist("ahk_id " MyGui.Hwnd) && WinActive("ahk_id " MyGui.Hwnd)) {
            wasVisible := true
            MyGui.Hide()
        }
    } catch {
        ; Window might not exist yet
    }
    
    Sleep(300) ; Longer delay for stability
    
    try {
        ; Get handle of currently active window
        focusedWin := WinGetID("A")
        if (focusedWin) {
            lastDetectedHwnd := focusedWin
            
            ; Show main window with proper theming
            MyGui.Show("w520 h450")
            
            ; FIXED: Uncomment and properly implement theme refresh after showing
            Sleep(100)  ; Allow window to fully render
            RefreshMainGUITheme()  ; Use the specific main GUI refresh function
            
            ; Additional comprehensive refresh to ensure all controls update
            SetTimer(DelayedMainWindowRefresh, -200)  ; Delayed refresh for stubborn controls
            
            UpdateShortcutsInfo()
        } else {
            ; Show window even if no valid window was captured
            MyGui.Show("w520 h450")
            Sleep(100)
            RefreshMainGUITheme()
            SetTimer(DelayedMainWindowRefresh, -200)
        }
    } catch as err {
        ; Show window even on error
        MyGui.Show("w520 h450")
        try {
            Sleep(100)
            RefreshMainGUITheme()
            SetTimer(DelayedMainWindowRefresh, -200)
        } catch {
            ; Theme refresh failed
        }
    }
}

/**
 * Enhanced DelayedMainWindowRefresh with specific Text control handling
 */
DelayedMainWindowRefresh() {
    try {
        ; Force a complete refresh of the main window
        RefreshGUITheme(MyGui, "Main Application Window")
        
        ; Apply enhanced text control theming specifically
        ForceTextControlsRetheming(MyGui)
        
        ; ADDED: Specific handling for header and label controls
        try {
            global ShortcutsText, CategoryFilter, AddAppButton, PreferencesButton, CloseButton, HelpButton, ExportButton
            global FilterLabel, ShortcutsHeader
            
            colors := GetThemeColors()
            isDark := IsWindowsDarkMode()
            
            ; Special attention to Text controls that are having issues
            if (IsObject(ShortcutsHeader)) {
                ApplyControlTheming(ShortcutsHeader, "Text", colors, isDark)
                ; Additional force refresh
                SetTimer(() => ForceTextControlRefresh(ShortcutsHeader), -100)
            }
            
            if (IsObject(FilterLabel)) {
                ApplyControlTheming(FilterLabel, "Text", colors, isDark)
                ; Additional force refresh
                SetTimer(() => ForceTextControlRefresh(FilterLabel), -150)
            }
            
            ; Re-apply theming to other controls
            if (IsObject(ShortcutsText))
                ApplyControlTheming(ShortcutsText, "Edit", colors, isDark)
            if (IsObject(CategoryFilter))
                ApplyControlTheming(CategoryFilter, "DropDownList", colors, isDark)
            if (IsObject(AddAppButton))
                ApplyControlTheming(AddAppButton, "Button", colors, isDark)
            if (IsObject(PreferencesButton))
                ApplyControlTheming(PreferencesButton, "Button", colors, isDark)
            if (IsObject(CloseButton))
                ApplyControlTheming(CloseButton, "Button", colors, isDark)
            if (IsObject(HelpButton))
                ApplyControlTheming(HelpButton, "Button", colors, isDark)
            if (IsObject(ExportButton))
                ApplyControlTheming(ExportButton, "Button", colors, isDark)
            
            ; Force complete window repaint
            MyGui.Redraw()
            
        } catch {
            ; Individual control theming failed, continue
        }
        
    } catch {
        ; Complete refresh failed
    }
}

/**
 * Enhanced UpdateShortcutsInfo to ensure proper theming after content updates
 */
UpdateShortcutsInfo(preserveFilter := false) {
    global lastDetectedHwnd, shortcuts_db, ShortcutsText, CategoryFilter
    
    if (lastDetectedHwnd = 0 || !WinExist("ahk_id " lastDetectedHwnd)) {
        ShortcutsText.Value := "No application captured or window no longer exists.`n`nPlease capture an application first."
        CategoryFilter.Delete()
        CategoryFilter.Add(["All Categories"])
        CategoryFilter.Choose(1)
        
        ; ADDED: Re-apply theming after content update
        try {
            colors := GetThemeColors()
            isDark := IsWindowsDarkMode()
            ApplyControlTheming(ShortcutsText, "Edit", colors, isDark)
            ApplyControlTheming(CategoryFilter, "DropDownList", colors, isDark)
        } catch {
            ; Theming failed, continue
        }
        return
    }
    
    ; Get application information
    processName := WinGetProcessName("ahk_id " lastDetectedHwnd)
    windowClass := WinGetClass("ahk_id " lastDetectedHwnd)
    appDisplayName := GetAppDisplayName(processName)
    
    ; Load shortcuts for this application
    shortcuts := GetShortcutsForApp(processName, windowClass)
    
    if (shortcuts && shortcuts.Length > 0) {
        if (!preserveFilter) {
            ; Rebuild category filter dropdown
            CategoryFilter.Delete()
            CategoryFilter.Add(["All Categories"])
            
            categorized := Map()
            
            for shortcut in shortcuts {
                category := shortcut.HasOwnProp("category") ? shortcut.category : "General"
                
                if !categorized.Has(category) {
                    categorized[category] := []
                    CategoryFilter.Add([category])
                }
                    
                categorized[category].Push(shortcut)
            }
            
            CategoryFilter.Choose(1)
            
            ; ADDED: Re-apply theming after rebuilding dropdown
            try {
                colors := GetThemeColors()
                isDark := IsWindowsDarkMode()
                ApplyControlTheming(CategoryFilter, "DropDownList", colors, isDark)
            } catch {
                ; Theming failed, continue
            }
        }
        
        selectedCategory := CategoryFilter.Text
        DisplayShortcutsWithFilter(shortcuts, selectedCategory)
    } else {
        ; No shortcuts found
        ShortcutsText.Value := "No shortcuts found in database for " appDisplayName ".`n`n"
            . "You can add this application to the database by clicking the 'Add to Database' button below."
        
        CategoryFilter.Delete()
        CategoryFilter.Add(["All Categories"])
        CategoryFilter.Choose(1)
        
        ; ADDED: Re-apply theming after content update
        try {
            colors := GetThemeColors()
            isDark := IsWindowsDarkMode()
            ApplyControlTheming(ShortcutsText, "Edit", colors, isDark)
            ApplyControlTheming(CategoryFilter, "DropDownList", colors, isDark)
        } catch {
            ; Theming failed, continue
        }
    }
}

/**
 * Update shortcuts display based on category filter selection
 * @return {Void}
 */
UpdateShortcutsWithFilter(*) {
    global CategoryFilter
    
    if (CategoryFilter.Text == "")
        CategoryFilter.Choose(1)
        
    UpdateShortcutsInfo(true)
}

/**
 * Display shortcuts with category filtering applied
 * @param {Array} shortcuts - Array of shortcut objects to display
 * @param {String} categoryFilter - Selected category filter ("All Categories" or specific category)
 * @return {Void}
 */
DisplayShortcutsWithFilter(shortcuts, categoryFilter) {
    global ShortcutsText, lastDetectedHwnd
    
    processName := WinGetProcessName("ahk_id " lastDetectedHwnd)
    appDisplayName := GetAppDisplayName(processName)
    
    categorized := Map()
    
    ; Organize shortcuts by category and apply filter
    for shortcut in shortcuts {
        category := shortcut.HasOwnProp("category") ? shortcut.category : "General"
        
        if (categoryFilter != "All Categories" && category != categoryFilter)
            continue
            
        if !categorized.Has(category)
            categorized[category] := []
            
        categorized[category].Push(shortcut)
    }
    
    ; Build display text
    shortcutInfo := "`nKeyboard Shortcuts for " appDisplayName ":`n`n"
    
    for category, categoryShortcuts in categorized {
        shortcutInfo .= "== " category " ==`n"
        
        for shortcut in categoryShortcuts {
            command := shortcut.HasOwnProp("command_name") ? shortcut.command_name : 
                      (shortcut.HasOwnProp("command") ? shortcut.command : "Unknown Command")
            keys := shortcut.HasOwnProp("shortcut_key") ? shortcut.shortcut_key : 
                   (shortcut.HasOwnProp("shortcut") ? shortcut.shortcut : "Unknown Shortcut")
            desc := shortcut.HasOwnProp("description") ? shortcut.description : 
                   (shortcut.HasOwnProp("desc") ? shortcut.desc : "")
            
            shortcutInfo .= command "`n"
            shortcutInfo .= keys "`n"
            if (desc != "")
                shortcutInfo .= desc "`n"
            
            shortcutInfo .= "`n"
        }
        
        shortcutInfo .= "`n"
    }
    
    if (categorized.Count = 0) {
        shortcutInfo .= "No shortcuts found for the selected category.`n"
    }
    
    ShortcutsText.Value := shortcutInfo
	; FIX: Re-apply the theme to the Edit control after changing its value
try {
    colors := GetThemeColors()
    isDark := IsWindowsDarkMode()
    ApplyControlTheming(ShortcutsText, "Edit", colors, isDark)
} catch {
    ; Silently ignore if theming fails here
}
}

/**
 * Get shortcuts for specific application from database
 * @param {String} processName - Process executable name
 * @param {String} windowClass - Window class name (for future matching)
 * @return {Array} Array of shortcut objects for the application
 */
GetShortcutsForApp(processName, windowClass) {
    global shortcuts_db
    
    ; Direct match by process name
    if shortcuts_db.Has(processName)
        return shortcuts_db[processName]
    
    ; Fuzzy match for partial process names
    for process, shortcuts in shortcuts_db {
        if InStr(processName, process) || InStr(process, processName)
            return shortcuts
    }
    
    return []
}

;==============================================================================
; APPLICATION LIFECYCLE MANAGEMENT WITH ENHANCED THEME SUPPORT
;==============================================================================


/**
 * Enhanced CloseApp function with proper cleanup and error handling
 * Handle application closing with proper data persistence and resource cleanup
 * @return {Void}
 */
CloseApp(*) {
    global shortcuts_db, DB_FILE, MyGui, activeGuis, activeFocusTimers
    
    ; NEW: Cleanup theme callbacks to prevent memory leaks
    try {
        CleanupThemeCallbacks()
    } catch {
        ; Ignore cleanup errors during shutdown
    }
    
    ; NEW: Cleanup any active GUI dialogs and their focus timers
    try {
        ; Stop all focus monitoring timers
        for guiHwnd, timerFunc in activeFocusTimers {
            try {
                SetTimer(timerFunc, 0)
            } catch {
                ; Timer might already be stopped
            }
        }
        activeFocusTimers.Clear()
        
        ; Close any remaining active dialogs
        for guiHwnd, closeCallback in activeGuis {
            try {
                if (WinExist("ahk_id " guiHwnd)) {
                    WinClose("ahk_id " guiHwnd)
                }
            } catch {
                ; Window might already be closed
            }
        }
        activeGuis.Clear()
        
    } catch {
        ; Ignore GUI cleanup errors during shutdown
    }
    
    ; NEW: Stop theme monitoring
    try {
        SetTimer(CheckThemeChange, 0)
        global intensiveThemeMonitoring
        intensiveThemeMonitoring := false
    } catch {
        ; Ignore timer cleanup errors
    }
    
    ; Save application data with enhanced error handling
    saveSuccess := false
    
    try {
        ; Attempt to save to SQLite database first
        if (InitializeSQLite()) {
            db := SQLite_Open(DB_FILE)
            
            if (db) {
                try {
                    SQLite_Exec(db, "BEGIN TRANSACTION;")
                    SQLite_Exec(db, "DELETE FROM shortcuts;")
                    
                    totalSaved := 0
                    
                    ; Save all shortcuts to database
                    for processName, shortcuts in shortcuts_db {
                        for _, shortcut in shortcuts {
                            success := AddShortcutToDatabase(db, 
                                processName, 
                                shortcut.HasOwnProp("command_name") ? shortcut.command_name : "", 
                                shortcut.HasOwnProp("shortcut_key") ? shortcut.shortcut_key : "", 
                                shortcut.HasOwnProp("category") ? shortcut.category : "General", 
                                shortcut.HasOwnProp("description") ? shortcut.description : ""
                            )
                            if (success) {
                                totalSaved++
                            }
                        }
                    }
                    
                    SQLite_Exec(db, "COMMIT;")
                    SQLite_Close(db)
                    
                    saveSuccess := true
                    
                    ; Also save backup file as redundancy
                    try {
                        SaveShortcutsToFile()
                    } catch {
                        ; Backup file save failed, but database save succeeded
                    }
                    
                } catch Error as dbError {
                    ; Database transaction failed, rollback and try backup file
                    try {
                        SQLite_Exec(db, "ROLLBACK;")
                        SQLite_Close(db)
                    } catch {
                        ; Rollback failed
                    }
                    
                    ; Fall through to backup file save
                }
            }
        }
        
        ; If database save failed, try backup file save
        if (!saveSuccess) {
            try {
                SaveShortcutsToFile()
                saveSuccess := true
            } catch Error as fileError {
                ; Both methods failed
                if (!quietMode) {
                    MsgBox("Warning: Could not save shortcuts before exit.`n`nDatabase Error: Unable to access SQLite`nFile Error: " fileError.Message, "Save Error", 48)
                }
            }
        }
        
    } catch Error as e {
        ; Complete save operation failed
        if (!quietMode) {
            MsgBox("Warning: Could not save shortcuts before exit.`nError: " e.Message, "Save Error", 48)
        }
    }
    
    ; NEW: Cleanup global hotkeys
    try {
        ; Disable all registered hotkeys
        try {
            Hotkey("`` & 2", "Off")
        } catch {
            ; Hotkey might not be registered
        }
        try {
            Hotkey("`` & 3", "Off")
        } catch {
            ; Hotkey might not be registered
        }
        try {
            Hotkey("Escape", "Off")
        } catch {
            ; Hotkey might not be registered
        }
        
        global globalHotkeysSuspended
        globalHotkeysSuspended := true
        
    } catch {
        ; Ignore hotkey cleanup errors
    }
    
    ; Hide main window safely
    try {
        if (IsObject(MyGui) && MyGui.HasOwnProp("Hwnd") && MyGui.Hwnd) {
            if (WinExist("ahk_id " MyGui.Hwnd)) {
                MyGui.Hide()
            }
        }
    } catch {
        ; Window might already be closed or destroyed
    }
    
    ; NEW: Final resource cleanup
    try {
        ; Clear global data structures
        shortcuts_db.Clear()
        global at_compat_db, at_usage_db
        if (IsObject(at_compat_db)) {
            at_compat_db.Clear()
        }
        if (IsObject(at_usage_db)) {
            at_usage_db.Clear()
        }
        
    } catch {
        ; Ignore data structure cleanup errors
    }
}


/**
 * Enhanced emergency shutdown function for critical errors
 * Force immediate application exit with minimal cleanup
 * @param {String} reason - Reason for emergency shutdown
 * @return {Void}
 */
EmergencyShutdown(reason := "Critical Error") {
    global quietMode
    
    ; Show error message if not in quiet mode
    if (!quietMode) {
        MsgBox("Emergency shutdown initiated: " reason "`n`nThe application will exit immediately.", "Emergency Shutdown", 16)
    }
    
    ; Minimal cleanup to prevent hanging
    try {
        CleanupThemeCallbacks()
    } catch {
        ; Ignore cleanup errors during emergency shutdown
    }
    
    try {
        ; Stop critical timers only
        SetTimer(CheckThemeChange, 0)
    } catch {
        ; Ignore timer errors
    }
    
    try {
        ; Clear message handlers to prevent further processing
        OnMessage(0x0100, "")
        OnMessage(0x001A, "")
        OnMessage(0x031A, "")
    } catch {
        ; Ignore handler cleanup errors
    }
    
    ; Force immediate exit without additional cleanup
    ExitApp(1)
}

/**
 * Display comprehensive help dialog with application information and enhanced theme checking
 * @return {Void}
 */
ShowHelpDialog(*) {
    global SCRIPT_NAME, SCRIPT_VERSION, MyGui

    ; Preserve original theme behavior
    ForceThemeCheck()
    try {
        HelpGui := CreateManagedDialog("+Owner" MyGui.Hwnd, SCRIPT_NAME " Help")
    } catch {
        HelpGui := Gui("+Owner" MyGui.Hwnd, SCRIPT_NAME " Help")
        ApplyDarkModeToGUI(HelpGui, "Help Dialog")
    }
    HelpGui.SetFont("s10")
    colors := GetThemeColors()
    isDark := IsWindowsDarkMode()

    ; You can edit this text later in-place
    helpContent := "
    ( Join`r`n LTrim
    KEYFINDER — WHAT IT IS
    ----------------------
    Keyfinder is a keyboard‑first utility that helps you find, learn, and copy
    keyboard shortcuts for Windows and supported apps.

    COMMON USES
    -----------
    • Look up shortcuts while working (stay in flow).
    • Copy shortcuts into email, docs, or training notes.
    • Verify the correct key combo when apps update.
    • Discover keyboard alternatives to mouse workflows.

    HOW TO USE
    ----------
    1) Search — type an action (e.g., “rename”, “bold”, “switch tab”) or a combo (e.g., “Ctrl+Shift+N”).
    2) Navigate — use ↑/↓ to move through results. Enter copies the selected shortcut. Esc closes.
    3) Settings — configure hotkeys, theme behavior, and other preferences (varies by build).
    )"

    ; ===== One scrollable text area =====
    HelpText := HelpGui.Add("Edit", "x10 y10 w580 h360 ReadOnly -Wrap VScroll -TabStop", helpContent)
    ApplyControlTheming(HelpText, "Edit", colors, isDark)

    ; ===== Company / phone line (static text) =====
    footer := HelpGui.Add("Text", "x10 y380 w580 +0x10", "EyeTech Analytics  |  Phone: (859) 212-3668")
    ApplyControlTheming(footer, "Text", colors, isDark)

    ; ===== Action buttons (links replaced with robust buttons) =====
    btnWeb := HelpGui.Add("Button", "x10 y406 w160", "Open Website")
    btnMail := HelpGui.Add("Button", "x180 y406 w160", "Email Support")
    CloseHelpButton := HelpGui.Add("Button", "Default x470 y406 w120", "&Close")

    btnWeb.OnEvent("Click", (*) => Run("https://eyetechanalytics.com"))
    btnMail.OnEvent("Click", (*) => Run("mailto:keyfinder@eyetechanalytics.com"))
    CloseHelpButton.OnEvent("Click", (*) => HelpGui.Destroy())

    ApplyControlTheming(btnWeb, "Button", colors, isDark)
    ApplyControlTheming(btnMail, "Button", colors, isDark)
    ApplyControlTheming(CloseHelpButton, "Button", colors, isDark)

    HelpGui.OnEvent("Escape", (*) => HelpGui.Destroy())
    HelpGui.Show()

    try ForceWindowToForeground(HelpGui.Hwnd)

    ; Preserve theme refresh sequence
    ForceThemeCheck()
    RefreshGUITheme(HelpGui, "Help Dialog")
    ForceTextControlsRetheming(HelpGui)
}

; Handler for Link clicks in the Help dialog
KF_Help_OpenLink(ctrl, info) {
    try {
        if IsObject(info) && info.HasProp("href") && info.href
            Run info.href
        else if info
            Run info
    } catch as _ {
        ; If the SysLink already handled it or Run fails, ignore silently.
    }
}


;==============================================================================
; MESSAGE HANDLER FOR ESCAPE KEY
;==============================================================================

/**
 * Custom message handler for Escape key in standard Windows dialogs and application windows
 * @param {Integer} wParam - Virtual key code (27 for Escape)
 * @param {Integer} lParam - Key data flags
 * @param {Integer} msg - Message type (WM_KEYDOWN)
 * @param {Integer} hwnd - Window handle receiving the message
 * @return {Void}
 */
WM_KEYDOWN(wParam, lParam, msg, hwnd) {
    global activeGuis, MyGui
    
    if (wParam = 27) { ; Escape key pressed
        ; First, check if this is one of our managed dialogs
        for guiHwnd, closeCallback in activeGuis {
            if (WinActive("ahk_id " guiHwnd)) {
                try {
                    closeCallback.Call()
                    return
                } catch {
                    ; Fallback to destroying the window
                    try {
                        WinClose("ahk_id " guiHwnd)
                    } catch {
                        ; Window might already be closed
                    }
                    return
                }
            }
        }
        
        ; Check if this is our main window
        if (WinActive("ahk_id " MyGui.Hwnd)) {
            try {
                MyGui.Hide()
                return
            } catch {
                ; Main window might not be available
            }
        }
        
        ; Handle standard Windows dialogs
        if WinActive("ahk_class #32770") { ; Standard Windows dialog
            title := WinGetTitle("A")
            
            ; Handle common dialog types with appropriate button clicks
            if (InStr(title, "Confirm") || InStr(title, "Delete") || InStr(title, "Warning") || 
                InStr(title, "Error") || InStr(title, "Question") || InStr(title, "Information")) {
                if ControlExist("Button2", "A")
                    ControlClick("Button2", "A")  ; Usually "Cancel" or "No"
                else if ControlExist("Button3", "A")
                    ControlClick("Button3", "A")  ; Alternative cancel button
                else if ControlExist("Button1", "A")
                    ControlClick("Button1", "A")  ; Last resort - usually "OK"
                else
                    ; Try to close the window directly
                    WinClose("A")
            } else {
                ; For other dialog types, try to close directly
                WinClose("A")
            }
        }
        
        ; For any other window type, try to close it if it belongs to our application
        activeWindow := WinGetID("A")
        try {
            processName := WinGetProcessName("ahk_id " activeWindow)
            currentProcessName := WinGetProcessName("ahk_id " A_ScriptHwnd)
            
            ; Only close windows from our own process
            if (processName = currentProcessName) {
                WinClose("ahk_id " activeWindow)
            }
        } catch {
            ; Ignore errors in process detection
        }
    }
}

;==============================================================================
; MISSING FUNCTION DEFINITIONS
;==============================================================================

/**
 * Handle uninstall button click from preferences dialog
 * @return {Void}
 */
StartUninstall(*) {
    HandleUninstall()
}

/**
 * Close active dialog or hide main window when Escape is pressed
 * @return {Void}
 */
CloseActiveDialogOrMainWindow() {
    global activeGuis, MyGui
    
    ; First, check if this is one of our managed dialogs
    for guiHwnd, closeCallback in activeGuis {
        if (WinActive("ahk_id " guiHwnd)) {
            try {
                closeCallback.Call()
                return
            } catch {
                ; Fallback to destroying the window
                try {
                    WinClose("ahk_id " guiHwnd)
                } catch {
                    ; Window might already be closed
                }
                return
            }
        }
    }
    
    ; Check if this is our main window
    try {
        if (IsObject(MyGui) && MyGui.Hwnd && WinActive("ahk_id " MyGui.Hwnd)) {
            MyGui.Hide()
            return
        }
    } catch {
        ; Main window might not be available
    }
    
    ; Handle any other active window from our process
    try {
        activeWindow := WinGetID("A")
        processName := WinGetProcessName("ahk_id " activeWindow)
        currentProcessName := WinGetProcessName("ahk_id " A_ScriptHwnd)
        
        ; Only close windows from our own process
        if (processName = currentProcessName) {
            WinClose("ahk_id " activeWindow)
        }
    } catch {
        ; Ignore errors in process detection
    }
}

/**
 * Convert human-readable hotkey format to AutoHotkey syntax
 * @param {String} humanReadable - Format like "Ctrl+Alt+2" or "`` & 2"
 * @return {String} AutoHotkey format like "^!2" or "`` & 2"
 */
ConvertToAutoHotkeyFormat(humanReadable) {
    if (humanReadable = "") {
        return ""
    }
    
    ; If it's already in backtick format, return as-is
    if (InStr(humanReadable, "`` & ")) {
        return humanReadable
    }
    
    ; Split by + to get components
    parts := StrSplit(humanReadable, "+")
    autoHotkeyFormat := ""
    regularKey := ""
    
    ; Process each part
    for part in parts {
        part := Trim(part)
        switch part {
            case "Ctrl":
                autoHotkeyFormat .= "^"
            case "Alt":
                autoHotkeyFormat .= "!"
            case "Shift":
                autoHotkeyFormat .= "+"
            case "Win":
                autoHotkeyFormat .= "#"
            default:
                ; This is the regular key
                regularKey := part
        }
    }
    
    ; Add the regular key at the end
    autoHotkeyFormat .= regularKey
    
    return autoHotkeyFormat
}


;==============================================================================
; SYSTEM TRAY INTEGRATION WITH ENHANCED THEME SUPPORT
;==============================================================================

/**
 * Initialize system tray icon and menu
 * @return {Void}
 */
InitializeSystemTray() {
    global SCRIPT_NAME, MyGui
    
    ; Create system tray menu
    TrayMenu := A_TrayMenu
    TrayMenu.Delete() ; Remove default menu items
    
    TrayMenu.Add("Show " SCRIPT_NAME, ShowMainWindow)
    TrayMenu.Add("Capture Window (`` & 2)", (*) => CaptureFocusedWindow())
    TrayMenu.Add("Search Shortcuts (`` & 3)", (*) => ShowSearchShortcutDialog())
    TrayMenu.Add()  ; Separator
    TrayMenu.Add("Settings", (*) => ShowPreferencesDialog())
    TrayMenu.Add("Help", (*) => ShowHelpDialog())
    TrayMenu.Add()  ; Separator
    TrayMenu.Add("Exit", ExitApplication)
    
    ; Set default action for double-click
    TrayMenu.Default := "Show " SCRIPT_NAME
    
    ; Set tray tooltip
    A_IconTip := SCRIPT_NAME " v" SCRIPT_VERSION " - Right-click for options"
}

/**
 * Show main window from system tray with enhanced theme checking
 * @return {Void}
 */
ShowMainWindow(*) {
    global MyGui
    
    ; COMPREHENSIVE THEME CHECK BEFORE SHOWING WINDOW
    ForceThemeCheck()
    
    try {
        if (!WinExist("ahk_id " MyGui.Hwnd)) {
            ; Window was destroyed, recreate it
            InitializeGUIWithThemeSupport()
        }
        
        MyGui.Show("w520 h450")
        
        ; Bring window to foreground
        ForceWindowToForeground(MyGui.Hwnd)
    } catch Error as e {
        MsgBox("Error showing main window: " e.Message, "Error", 16)
    }
}

/**
 * Enhanced ExitApplication function with comprehensive cleanup
 * Exit application completely with proper resource cleanup
 * @return {Void}
 */
ExitApplication(*) {
    global SCRIPT_NAME
    
    ; NEW: Cleanup callbacks and resources before exit
    try {
        CleanupThemeCallbacks()
    } catch {
        ; Ignore cleanup errors
    }
    
    ; NEW: Stop all active timers to prevent issues during shutdown
    try {
        ; Stop theme monitoring
        SetTimer(CheckThemeChange, 0)
        
        ; Stop any delayed refresh timers
        SetTimer(DelayedThemeRefresh, 0)
        SetTimer(ScheduleThemeRefreshCallback, 0)
        SetTimer(DelayedMainWindowRefresh, 0)
        SetTimer(FinalMainWindowRefresh, 0)
        SetTimer(AdditionalTextControlRefresh, 0)
        
    } catch {
        ; Ignore timer cleanup errors
    }
    
    ; NEW: Clear Windows message handlers
    try {
        OnMessage(0x0100, "")  ; Clear WM_KEYDOWN handler
        OnMessage(0x001A, "")  ; Clear WM_WININICHANGE handler
        OnMessage(0x031A, "")  ; Clear WM_THEMECHANGED handler
        OnMessage(0x02B1, "")  ; Clear WM_WTSSESSION_CHANGE handler
    } catch {
        ; Ignore message handler cleanup errors
    }
    
    ; NEW: Force cleanup of any remaining GUI objects
    try {
        global activeGuis, activeFocusTimers
        
        ; Emergency cleanup of any remaining dialogs
        if (IsObject(activeGuis)) {
            for guiHwnd, closeCallback in activeGuis {
                try {
                    if (WinExist("ahk_id " guiHwnd)) {
                        ; Force close without callbacks to prevent hanging
                        DllCall("DestroyWindow", "Ptr", guiHwnd)
                    }
                } catch {
                    ; Ignore individual window cleanup errors
                }
            }
            activeGuis.Clear()
        }
        
        ; Stop all focus timers
        if (IsObject(activeFocusTimers)) {
            for guiHwnd, timerFunc in activeFocusTimers {
                try {
                    SetTimer(timerFunc, 0)
                } catch {
                    ; Timer might already be stopped
                }
            }
            activeFocusTimers.Clear()
        }
        
    } catch {
        ; Ignore GUI cleanup errors
    }
    
    ; Save data before exiting using the enhanced CloseApp function
    try {
        CloseApp()
    } catch Error as e {
        ; Even if save fails, continue with exit
        if (!quietMode) {
            MsgBox("Warning: Error during shutdown: " e.Message "`n`nThe application will exit anyway.", "Shutdown Warning", 48)
        }
    }
    
    ; NEW: Final system cleanup
    try {
        ; Clear any remaining system tray menu
        try {
            A_TrayMenu.Delete()
        } catch {
            ; Tray menu cleanup failed
        }
        
        ; Reset tray tooltip
        try {
            A_IconTip := "AutoHotkey Script"
        } catch {
            ; Tooltip reset failed
        }
        
    } catch {
        ; Ignore system cleanup errors
    }
    
    ; NEW: Brief delay to allow cleanup to complete
    Sleep(100)
    
    ; Exit the application
    ExitApp(0)
}

;==============================================================================
; ENHANCED SHORTCUT SEARCH FUNCTIONALITY WITH COMPREHENSIVE THEME CHECKING
;==============================================================================

/**
 * Display keystroke search dialog with enhanced key capture and comprehensive theme checking
 * @return {Void}
 */
ShowSearchShortcutDialog(*) {
    global shortcuts_db, DB_FILE
    
    ; COMPREHENSIVE THEME CHECK BEFORE DIALOG CREATION
    ForceThemeCheck()
    
    try {
        SearchGui := CreateManagedDialog("", "Search Shortcuts by Keystroke")
    } catch {
        SearchGui := Gui("", "Search Shortcuts by Keystroke")
        ApplyDarkModeToGUI(SearchGui, "Search Dialog")
    }
    SearchGui.SetFont("s10")
    colors := GetThemeColors()
    isDark := IsWindowsDarkMode()
    
    instructionText := SearchGui.Add("Text", "x10 y10 w380", "Press the keystroke combination you want to search for:")
	ApplyControlTheming(instructionText, "Text", colors, isDark)
    
    keystrokeLabel := SearchGui.Add("Text", "x10 y40 w100", "Keystroke:")
    ApplyControlTheming(keystrokeLabel, "Text", colors, isDark)
    
    KeystrokeEdit := SearchGui.Add("Edit", "x120 y40 w250 ReadOnly", "")
    ApplyControlTheming(KeystrokeEdit, "Edit", colors, isDark)
    KeystrokeEdit.Focus()
    
    ; Key capture state variables
    pressedKeys := Map()
    capturedKeys := Map()
    keystrokeReady := false
    isCapturing := false
    hasCaptured := false
    
    ; Control buttons
    SearchButton := SearchGui.Add("Button", "Default x100 y80 w100", "Search")
    SearchButton.Enabled := false
    ApplyControlTheming(SearchButton, "Button", colors, isDark)
    
    ClearButton := SearchGui.Add("Button", "x210 y80 w80", "Clear")
    ApplyControlTheming(ClearButton, "Button", colors, isDark)
    
    CancelButton := SearchGui.Add("Button", "x300 y80 w80", "Cancel")
    ApplyControlTheming(CancelButton, "Button", colors, isDark)
    
    ; Button event handlers
    SearchButton.OnEvent("Click", SearchShortcuts)
    ClearButton.OnEvent("Click", ClearKeystroke)
    CancelButton.OnEvent("Click", (*) => SearchGui.Destroy())
    SearchGui.OnEvent("Close", (*) => SearchGui.Destroy())
    
    ; Hotkey setup for this dialog
    HotIfWinActive("ahk_id " SearchGui.Hwnd)
    
    Hotkey "Enter", ExecuteSearch
    Hotkey "NumpadEnter", ExecuteSearch
    ; Ensure Escape works even during key capture
    Hotkey "Escape", (*) => SearchGui.Destroy()
    
    RegisterHotkeys()
    
    ; Start capturing after short delay
    SetTimer(StartCapturing, -300)
    
    StartCapturing() {
        isCapturing := true
    }
    
    ExecuteSearch(*) {
        if (SearchButton.Enabled && keystrokeReady)
            SearchShortcuts()
    }
    
    RegisterHotkeys() {
        ; Register modifier keys
        RegisterKeyPair("LCtrl")
        RegisterKeyPair("RCtrl")
        RegisterKeyPair("LAlt")
        RegisterKeyPair("RAlt")
        RegisterKeyPair("LShift")
        RegisterKeyPair("RShift")
        RegisterKeyPair("LWin")
        RegisterKeyPair("RWin")
        
        ; Register alphabet keys
        alphabet := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Loop Parse, alphabet
            RegisterKeyPair(A_LoopField)
            
        ; Register number keys
        Loop 10
            RegisterKeyPair(A_Index-1)
            
        ; Register function keys
        Loop 12
            RegisterKeyPair("F" A_Index)
            
        ; Register other common keys
        otherKeys := ["Tab", "Space", "Delete", "Home", "End", 
                     "PgUp", "PgDn", "Insert", "NumpadAdd", "NumpadSub", 
                     "NumpadMult", "NumpadDiv", "Left", "Right", "Up", "Down"]
        For key in otherKeys
            RegisterKeyPair(key)
    }
 
    RegisterKeyPair(key) {
        Hotkey("~*" key, KeyDown.Bind(key))
        Hotkey("~*" key " up", KeyUp.Bind(key, false))
    }
    
    KeyDown(key, *) {
        if (!isCapturing)
            return
        
        if (hasCaptured && pressedKeys.Count = 0) {
            hasCaptured := false
            capturedKeys := Map()
        }
        
        pressedKeys[key] := true
        UpdateKeystrokeDisplay(pressedKeys)
        
        keystrokeReady := true
        SearchButton.Enabled := true
    }
    
    KeyUp(key, isAllKeysUp, *) {
        if (!isCapturing)
            return
            
        if (pressedKeys.Has(key))
            pressedKeys.Delete(key)
        
        if (pressedKeys.Count = 0 && SearchButton.Enabled) {
            hasCaptured := true
            capturedKeys := Map()
        }
    }
    
    UpdateKeystrokeDisplay(keysMap) {
        if (hasCaptured && keysMap.Count = 0)
            return
            
        modifiers := []
        regularKeys := []
        
        ; Categorize keys into modifiers and regular keys
        for key, _ in keysMap {
            if (key = "LCtrl" || key = "RCtrl")
                AddToArrayIfMissing(modifiers, "Ctrl")
            else if (key = "LAlt" || key = "RAlt")
                AddToArrayIfMissing(modifiers, "Alt")
            else if (key = "LShift" || key = "RShift")
                AddToArrayIfMissing(modifiers, "Shift")
            else if (key = "LWin" || key = "RWin")
                AddToArrayIfMissing(modifiers, "Win")
            else
                AddToArrayIfMissing(regularKeys, key)
        }
        
        ; Build keystroke text representation
        keystrokeText := ""
        
        if (modifiers.Length > 0)
            keystrokeText := Join(modifiers, "+")
            
        if (regularKeys.Length > 0) {
            if (keystrokeText != "")
                keystrokeText .= "+"
                
            keystrokeText .= Join(regularKeys, "+")
        }
        
        KeystrokeEdit.Value := keystrokeText

        ; <<< FIX: Re-apply theming to the Edit control after updating its value
        ApplyControlTheming(KeystrokeEdit, "Edit",colors, isDark)
    }
    
  
   
    ClearKeystroke(*) {
        pressedKeys := Map()
        capturedKeys := Map()
        hasCaptured := false
        keystrokeReady := false
        KeystrokeEdit.Value := ""
        SearchButton.Enabled := false
        
        ; <<< FIX: Re-apply theming to the Edit control after clearing its value
        ApplyControlTheming(KeystrokeEdit, "Edit", colors, isDark)
    }
    
    SearchShortcuts(*) {
        keystroke := KeystrokeEdit.Value
        
        if (keystroke = "") {
            MsgBox("Please enter a valid keystroke combination.", "Input Required", 48)
            return
        }
        
        matches := []
        
        ; Search in memory database
        for processName, shortcuts in shortcuts_db {
            for shortcut in shortcuts {
                if (shortcut.HasOwnProp("shortcut_key") && StrCompare(shortcut.shortcut_key, keystroke, true) = 0) {
                    matches.Push({
                        process_name: processName,
                        command_name: shortcut.HasOwnProp("command_name") ? shortcut.command_name : "",
                        category: shortcut.HasOwnProp("category") ? shortcut.category : "General",
                        description: shortcut.HasOwnProp("description") ? shortcut.description : ""
                    })
                }
            }
        }
        
        ; Search in SQLite database if available
        if (InitializeSQLite()) {
            try {
                db := SQLite_Open(DB_FILE)
                if db {
                    sql := "SELECT process_name, command_name, category, description FROM shortcuts WHERE LOWER(shortcut_key) = LOWER(?);"
                    results := SQLite_Query(db, sql, keystroke)
                    
                    for row in results {
                        ; Check for duplicates from memory search
                        isDuplicate := false
                        for match in matches {
                            if (match.process_name = row.process_name && match.command_name = row.command_name) {
                                isDuplicate := true
                                break
                            }
                        }
                        
                        if (!isDuplicate) {
                            matches.Push({
                                process_name: row.process_name,
                                command_name: row.command_name,
                                category: row.category || "General",
                                description: row.description || ""
                            })
                        }
                    }
                    
                    SQLite_Close(db)
                }
            } catch Error as e {
                ; Continue with memory results only
            }
        }
        
        HotIf()
        SearchGui.Destroy()
        ShowSearchResults(keystroke, matches)
    }
    
    RegisterGuiForEscapeHandling(SearchGui, (*) => SearchGui.Destroy())
    
    SearchGui.Show("w400 h120")
    ForceWindowToForeground(SearchGui.Hwnd)
    
    ; Final theme consistency check after showing
    ForceThemeCheck()
    RefreshGUITheme(SearchGui, "Search Dialog")
    
    ; Apply enhanced text control theming
    ForceTextControlsRetheming(SearchGui)
    
    HotIf()
}
/**
 * Display search results for keystroke matches with enhanced theme checking
 * @param {String} keystroke - The searched keystroke combination
 * @param {Array} matches - Array of matching shortcut objects
 * @return {Void}
 */
ShowSearchResults(keystroke, matches) {
    ; COMPREHENSIVE THEME CHECK BEFORE DIALOG CREATION
    ForceThemeCheck()
    
    try {
        ResultsGui := CreateManagedDialog("+Resize +MinSize500x350", "Shortcut Search Results")
    } catch {
        ResultsGui := Gui("+Resize +MinSize500x350", "Shortcut Search Results")
        ApplyDarkModeToGUI(ResultsGui, "Search Results")
    }
    ResultsGui.SetFont("s10")
    colors := GetThemeColors()
    isDark := IsWindowsDarkMode()
    
    HeaderText := ResultsGui.Add("Text", "x10 y10 w680", "Applications using the keystroke: " keystroke)
    ApplyControlTheming(HeaderText, "Text", colors, isDark)
    
    ; Create results ListView with all columns
    LV := ResultsGui.Add("ListView", "x10 y40 w680 h300 Grid", ["Application", "Command", "Category", "Description"])
    ApplyControlTheming(LV, "ListView", colors, isDark)
    
    LV.ModifyCol(1, 180)    ; Application
    LV.ModifyCol(2, 160)    ; Command  
    LV.ModifyCol(3, 100)    ; Category
    LV.ModifyCol(4, 230)    ; Description
    
    ; Populate results
    if (matches.Length > 0) {
        for match in matches {
            appDisplayName := GetAppDisplayName(match.process_name)
            description := match.HasOwnProp("description") ? match.description : ""
            LV.Add(, appDisplayName, match.command_name, match.category, description)
        }
    } else {
        LV.Add(, "No applications found using this keystroke", "", "", "")
    }
    
    CloseButton := ResultsGui.Add("Button", "x300 y350 w100", "Close")
    ApplyControlTheming(CloseButton, "Button", colors, isDark)
    CloseButton.OnEvent("Click", (*) => ResultsGui.Destroy())
    ResultsGui.OnEvent("Close", (*) => ResultsGui.Destroy())
    
    ; Search results resize handler
    OnSearchResultsResize(GuiObj, MinMax, Width, Height) {
        ; Enforce minimum window size
        minWidth := 500
        minHeight := 350
        
        if (Width < minWidth)
            Width := minWidth
        if (Height < minHeight)
            Height := minHeight
        
        ; Calculate content dimensions
        contentWidth := Width - 20
        
        try {
            ; Resize header text
            if (IsObject(HeaderText)) {
                HeaderText.Move(10, 10, contentWidth, 20)
            }
            
            ; Calculate ListView height (reserve space for close button)
            listHeight := Height - 100  ; 40 (header + margin) + 60 (button area) = 100
            if (listHeight < 200)  ; Minimum list height
                listHeight := 200
            
            ; Resize ListView
            if (IsObject(LV)) {
                LV.Move(10, 40, contentWidth, listHeight)
                
                ; Adjust column widths proportionally
                totalWidth := contentWidth - 20  ; Account for scrollbar space
                LV.ModifyCol(1, totalWidth * 0.25)    ; Application: 25%
                LV.ModifyCol(2, totalWidth * 0.22)    ; Command: 22%
                LV.ModifyCol(3, totalWidth * 0.15)    ; Category: 15%
                LV.ModifyCol(4, totalWidth * 0.38)    ; Description: 38%
            }
            
            ; Position Close button at bottom center
            buttonY := 40 + listHeight + 10
            buttonX := (Width - 100) / 2
            
            if (IsObject(CloseButton)) {
                CloseButton.Move(buttonX, buttonY, 100, 25)
            }
            
            ; Force visual refresh
            ResultsGui.Redraw()
            
        } catch Error as e {
            ; Silently handle control access errors during resize
        }
    }
    
    ResultsGui.OnEvent("Size", OnSearchResultsResize)
    
    RegisterGuiForEscapeHandling(ResultsGui, (*) => ResultsGui.Destroy())
    
    ; Show window with proper dimensions for description column
    ResultsGui.Show("w700 h400")
    ForceWindowToForeground(ResultsGui.Hwnd)
    
    ; Final theme consistency check after showing
    ForceThemeCheck()
    RefreshGUITheme(ResultsGui, "Search Results")
    
    ; Apply enhanced text control theming
    ForceTextControlsRetheming(ResultsGui)
}

;==============================================================================
; ENHANCED SHORTCUT MANAGEMENT DIALOGS WITH COMPREHENSIVE THEME CHECKING
;==============================================================================

/**
 * Sort array in descending order using bubble sort algorithm
 * @param {Array} arr - Array to sort in descending order
 * @return {Array} New array sorted in descending order
 */

/**
 * Sort array in descending order using the efficient built-in method.
 * @param {Array} arr - Array to sort.
 * @return {Array} The same array, now sorted.
 */
SortArrayDesc(arr) {
    arr.Sort("D") ; The "D" flag tells it to sort in descending order.
    return arr
}

/**
 * Display comprehensive application shortcut management dialog with enhanced theme checking and proper resizing
 * @return {Void}
 */
ShowAddAppDialog(*) {
    global lastDetectedHwnd, MyGui
    
    ; COMPREHENSIVE THEME CHECK BEFORE DIALOG CREATION
    ForceThemeCheck()
    
    if (lastDetectedHwnd = 0 || !WinExist("ahk_id " lastDetectedHwnd)) {
        MsgBox("Please capture an application first using `` & 2.", "No Application Detected", 48)
        return
    }
    
    processName := WinGetProcessName("ahk_id " lastDetectedHwnd)
    windowClass := WinGetClass("ahk_id " lastDetectedHwnd)
    appDisplayName := GetAppDisplayName(processName)
    
    try {
        ; Set minimum size to initial display size to prevent artifacting
        AddAppGui := CreateManagedDialog("+Owner" MyGui.Hwnd " +Resize +MinSize620x550", "Add to Database - " appDisplayName)
    } catch {
        AddAppGui := Gui("+Owner" MyGui.Hwnd " +Resize +MinSize620x550", "Add to Database - " appDisplayName)
        ; Apply theme manually if CreateManagedDialog failed
        ApplyDarkModeToGUI(AddAppGui, "Shortcut Management Dialog")
    }
    AddAppGui.SetFont("s10")
    colors := GetThemeColors()
    isDark := IsWindowsDarkMode()
    
    AddAppGui.OnEvent("Close", (*) => AddAppGui.Destroy())
    
    ; Dialog header with better spacing
    TitleText := AddAppGui.Add("Text", "x15 y15 w570 center", "Managing Shortcuts for Application: " appDisplayName)
    ApplyControlTheming(TitleText, "Text", colors, isDark)
    
    SubtitleText := AddAppGui.Add("Text", "x15 y40 w570 center", "Window Class: " windowClass)
    ApplyControlTheming(SubtitleText, "Text", colors, isDark)
    
    AccessInstructions := AddAppGui.Add("Text", "x15 y65 w570 center", "Select a shortcut to edit, or click 'Add New' to create one")
    AccessInstructions.SetFont("italic s9")
    ApplyControlTheming(AccessInstructions, "Text", colors, isDark)
    
    ; Shortcuts list section with enhanced spacing
    ShortcutsListLabel := AddAppGui.Add("Text", "x15 y95 w570", "Shortcuts List (select a shortcut to edit)")
    ShortcutsListLabel.SetFont("bold")
    ApplyControlTheming(ShortcutsListLabel, "Text", colors, isDark)
    
    ; Increased ListView height for better visibility
    LV := AddAppGui.Add("ListView", "x15 y120 w570 h140 Grid AltSubmit", ["Category", "Command", "Shortcut", "Description"])
    ApplyControlTheming(LV, "ListView", colors, isDark)
    
    LV.ModifyCol(1, 110, "Category")
    LV.ModifyCol(2, 140, "Command")
    LV.ModifyCol(3, 100, "Shortcut")
    LV.ModifyCol(4, 200, "Description")
    
    LVHelper := AddAppGui.Add("Text", "x15 y270 w570", "Double-click a shortcut to edit, or use buttons below")
    LVHelper.SetFont("italic s9")
    ApplyControlTheming(LVHelper, "Text", colors, isDark)
    
    ; Load existing shortcuts for this application
    existingShortcuts := GetShortcutsForApp(processName, windowClass)
    
    if (existingShortcuts && existingShortcuts.Length > 0) {
        for shortcut in existingShortcuts {
            category := shortcut.HasOwnProp("category") ? shortcut.category : "General"
            command := shortcut.HasOwnProp("command_name") ? shortcut.command_name : 
                       (shortcut.HasOwnProp("command") ? shortcut.command : "Unknown Command")
            keys := shortcut.HasOwnProp("shortcut_key") ? shortcut.shortcut_key : 
                   (shortcut.HasOwnProp("shortcut") ? shortcut.shortcut : "Unknown Shortcut")
            desc := shortcut.HasOwnProp("description") ? shortcut.description : 
                   (shortcut.HasOwnProp("desc") ? shortcut.desc : "")
            
            LV.Add(, category, command, keys, desc)
        }
    }
    
    ; Edit section with better spacing and larger area
    EditSectionGB := AddAppGui.Add("GroupBox", "x15 y295 w570 h180", "Edit Selected Shortcut")
    ApplyControlTheming(EditSectionGB, "GroupBox", colors, isDark)
    
    ; Category row with better spacing
    CategoryTextLabel := AddAppGui.Add("Text", "x25 y325 w80", "&Category:")
    ApplyControlTheming(CategoryTextLabel, "Text", colors, isDark)
    
    CategoryEdit := AddAppGui.Add("Edit", "x110 y325 w180 vCategoryEdit")
    ApplyControlTheming(CategoryEdit, "Edit", colors, isDark)
    
    CommonLabel := AddAppGui.Add("Text", "x300 y325 w80", "Co&mmon:")
    ApplyControlTheming(CommonLabel, "Text", colors, isDark)
    
    CategoryList := AddAppGui.Add("DropDownList", "x385 y325 w180", ["General", "File Operations", "Text Formatting", "Navigation", "Editing", "View Controls", "Accessibility", "Tables", "NVDA", "JAWS"])
    ApplyControlTheming(CategoryList, "DropDownList", colors, isDark)
    CategoryList.OnEvent("Change", (*) => CategoryEdit.Value := CategoryList.Text)
    
    ; Command row
    CmdTextLabel := AddAppGui.Add("Text", "x25 y355 w80", "Comman&d:")
    ApplyControlTheming(CmdTextLabel, "Text", colors, isDark)
    
    CommandEdit := AddAppGui.Add("Edit", "x110 y355 w455 vCommandEdit")
    ApplyControlTheming(CommandEdit, "Edit", colors, isDark)
    
    ; Shortcut row
    ShortcutTextLabel := AddAppGui.Add("Text", "x25 y385 w80", "&Shortcut:")
    ApplyControlTheming(ShortcutTextLabel, "Text", colors, isDark)
    
    ShortcutEdit := AddAppGui.Add("Edit", "x110 y385 w365 vShortcutEdit ReadOnly", "")
    ApplyControlTheming(ShortcutEdit, "Edit", colors, isDark)
    
    ClearShortcutBtn := AddAppGui.Add("Button", "x485 y385 w80", "&Clear")
    ApplyControlTheming(ClearShortcutBtn, "Button", colors, isDark)
    
    ShortcutHelp := AddAppGui.Add("Text", "x110 y410 w455", "Focus this field and press the key combination you want to capture")
    ApplyControlTheming(ShortcutHelp, "Text", colors, isDark)
    
    ; Description row with more height
    DescTextLabel := AddAppGui.Add("Text", "x25 y435 w80", "&Description:")
    ApplyControlTheming(DescTextLabel, "Text", colors, isDark)
    
    DescriptionEdit := AddAppGui.Add("Edit", "x110 y435 w455 h35 Multi VScroll vDescriptionEdit")
    ApplyControlTheming(DescriptionEdit, "Edit", colors, isDark)
    
    ; Action buttons with better spacing
    AddNewButton := AddAppGui.Add("Button", "x25 y490 w100", "Add &New")
    ApplyControlTheming(AddNewButton, "Button", colors, isDark)
    
    UpdateButton := AddAppGui.Add("Button", "x135 y490 w100", "&Update")
    UpdateButton.Enabled := false
    ApplyControlTheming(UpdateButton, "Button", colors, isDark)
    
    BulkAddButton := AddAppGui.Add("Button", "x245 y490 w100", "&Bulk Add")
    ApplyControlTheming(BulkAddButton, "Button", colors, isDark)
    
    DeleteButton := AddAppGui.Add("Button", "x355 y490 w100", "&Delete")
    ApplyControlTheming(DeleteButton, "Button", colors, isDark)
    
    ; Bottom row buttons
    SaveButton := AddAppGui.Add("Button", "Default x200 y520 w100", "&Save All")
    ApplyControlTheming(SaveButton, "Button", colors, isDark)
    
    CancelButton := AddAppGui.Add("Button", "x310 y520 w100", "C&ancel")
    ApplyControlTheming(CancelButton, "Button", colors, isDark)
    
    ; ADD RESIZE EVENT HANDLER
    AddAppGui.OnEvent("Size", OnAddAppDialogResize)
    
    ; Resize handler function
    OnAddAppDialogResize(GuiObj, MinMax, Width, Height) {
        ; Use initial display size as minimum to prevent artifacting
        minWidth := 620   ; Initial display width
        minHeight := 550  ; Initial display height
        
        ; Enforce minimum sizes (prevent shrinking below initial size)
        needsResize := false
        if (Width < minWidth) {
            Width := minWidth
            needsResize := true
        }
        if (Height < minHeight) {
            Height := minHeight
            needsResize := true
        }
        
        ; If we needed to enforce minimums, resize the window
        if (needsResize) {
            GuiObj.Move(, , Width, Height)
            return  ; Exit early to prevent double-processing
        }
        
        ; Disable window updates during resize to prevent flickering
        DllCall("LockWindowUpdate", "Ptr", GuiObj.Hwnd)
        
        try {
            ; Calculate content dimensions with margins
            contentWidth := Width - 30  ; 15px margins on each side
            
            ; Resize header elements
            if (IsObject(TitleText)) {
                TitleText.Move(15, 15, contentWidth, 20)
            }
            if (IsObject(SubtitleText)) {
                SubtitleText.Move(15, 40, contentWidth, 20)
            }
            if (IsObject(AccessInstructions)) {
                AccessInstructions.Move(15, 65, contentWidth, 20)
            }
            
            ; Resize shortcuts list section
            if (IsObject(ShortcutsListLabel)) {
                ShortcutsListLabel.Move(15, 95, contentWidth, 20)
            }
            
            ; Calculate space allocation for flexible ListView
            headerHeight := 110          ; Header area
            listHelperHeight := 25       ; ListView helper text + margin
            editSectionHeight := 180     ; Fixed edit section height
            actionButtonsHeight := 35    ; Action buttons + margin
            bottomButtonsHeight := 55    ; Bottom buttons + margin
            
            fixedSpaceUsed := headerHeight + listHelperHeight + editSectionHeight + actionButtonsHeight + bottomButtonsHeight
            availableForListView := Height - fixedSpaceUsed
            listHeight := Max(140, availableForListView)  ; Minimum 140px for ListView
            
            if (IsObject(LV)) {
                LV.Move(15, 120, contentWidth, listHeight)
                
                ; Adjust column widths proportionally
                totalColWidth := contentWidth - 25  ; Account for scrollbar
                LV.ModifyCol(1, totalColWidth * 0.18)    ; Category: 18%
                LV.ModifyCol(2, totalColWidth * 0.25)    ; Command: 25%
                LV.ModifyCol(3, totalColWidth * 0.18)    ; Shortcut: 18%
                LV.ModifyCol(4, totalColWidth * 0.39)    ; Description: 39%
            }
            
            ; Position ListView helper text immediately after ListView
            listBottomY := 120 + listHeight + 5
            if (IsObject(LVHelper)) {
                LVHelper.Move(15, listBottomY, contentWidth, 20)
            }
            
            ; Position edit section after helper text with margin
            editSectionY := listBottomY + 25
            if (IsObject(EditSectionGB)) {
                EditSectionGB.Move(15, editSectionY, contentWidth, editSectionHeight)
            }
            
            ; Position edit controls within the GroupBox (relative to edit section)
            editBaseY := editSectionY + 30
            
            ; Category row
            if (IsObject(CategoryTextLabel)) {
                CategoryTextLabel.Move(25, editBaseY, 80, 16)
            }
            if (IsObject(CategoryEdit)) {
                CategoryEdit.Move(110, editBaseY, 180, 21)
            }
            if (IsObject(CommonLabel)) {
                CommonLabel.Move(300, editBaseY, 80, 16)
            }
            if (IsObject(CategoryList)) {
                CategoryList.Move(contentWidth - 195, editBaseY, 180, 21)
            }
            
            ; Command row
            cmdY := editBaseY + 30
            if (IsObject(CmdTextLabel)) {
                CmdTextLabel.Move(25, cmdY, 80, 16)
            }
            if (IsObject(CommandEdit)) {
                CommandEdit.Move(110, cmdY, contentWidth - 140, 21)
            }
            
            ; Shortcut row
            shortcutY := editBaseY + 60
            if (IsObject(ShortcutTextLabel)) {
                ShortcutTextLabel.Move(25, shortcutY, 80, 16)
            }
            if (IsObject(ShortcutEdit)) {
                ShortcutEdit.Move(110, shortcutY, contentWidth - 225, 21)
            }
            if (IsObject(ClearShortcutBtn)) {
                ClearShortcutBtn.Move(contentWidth - 95, shortcutY, 80, 23)
            }
            if (IsObject(ShortcutHelp)) {
                ShortcutHelp.Move(110, shortcutY + 25, contentWidth - 140, 16)
            }
            
            ; Description row
            descY := editBaseY + 110
            if (IsObject(DescTextLabel)) {
                DescTextLabel.Move(25, descY, 80, 16)
            }
            if (IsObject(DescriptionEdit)) {
                DescriptionEdit.Move(110, descY, contentWidth - 140, 35)
            }
            
            ; Position action buttons immediately after edit section
            actionButtonY := editSectionY + editSectionHeight + 10
            
            if (IsObject(AddNewButton)) {
                AddNewButton.Move(25, actionButtonY, 100, 25)
            }
            if (IsObject(UpdateButton)) {
                UpdateButton.Move(135, actionButtonY, 100, 25)
            }
            if (IsObject(BulkAddButton)) {
                BulkAddButton.Move(245, actionButtonY, 100, 25)
            }
            if (IsObject(DeleteButton)) {
                DeleteButton.Move(355, actionButtonY, 100, 25)
            }
            
            ; Position bottom row buttons after action buttons
            bottomButtonY := actionButtonY + 35
            totalBottomButtonWidth := 210  ; 100 + 10 (gap) + 100
            startX := (Width - totalBottomButtonWidth) / 2
            
            if (IsObject(SaveButton)) {
                SaveButton.Move(startX, bottomButtonY, 100, 25)
            }
            if (IsObject(CancelButton)) {
                CancelButton.Move(startX + 110, bottomButtonY, 100, 25)
            }
            
        } catch Error as e {
            ; Silently handle control access errors during resize
        } finally {
            ; Re-enable window updates and force complete redraw
            DllCall("LockWindowUpdate", "Ptr", 0)
            
            ; Force comprehensive redraw to prevent artifacts
            try {
                ; Multiple redraw methods for better reliability
                GuiObj.Redraw()
                DllCall("InvalidateRect", "Ptr", GuiObj.Hwnd, "Ptr", 0, "Int", 1)
                DllCall("UpdateWindow", "Ptr", GuiObj.Hwnd)
                DllCall("RedrawWindow", "Ptr", GuiObj.Hwnd, "Ptr", 0, "Ptr", 0, "UInt", 0x0001 | 0x0004 | 0x0010)
            } catch {
                ; Fallback to basic redraw if advanced methods fail
                try {
                    GuiObj.Redraw()
                } catch {
                    ; Even basic redraw failed, continue
                }
            }
        }
    }
    
    ; [REST OF THE FUNCTION REMAINS THE SAME - all the event handlers, key capture, etc.]
    ; Set up key capture system and event handlers
    shortcutPressedKeys := Map()
    shortcutKeystrokeReady := false
    shortcutIsCapturing := false
    shortcutHasCaptured := false
    shortcutHotkeysRegistered := false
    
    ; State tracking for editing
    selectedRow := 0
    isEditing := false
    
    ; Define helper functions
    AddUniqueToArray(arr, item) {
        for existingItem in arr {
            if (existingItem = item)
                return
        }
        arr.Push(item)
    }
    
    ClearEditFields() {
        CategoryEdit.Value := ""
        CommandEdit.Value := ""
        ShortcutEdit.Value := ""
        DescriptionEdit.Value := ""
    }
    
    ClearShortcutField(*) {
        ShortcutEdit.Value := ""
        shortcutPressedKeys := Map()
        shortcutHasCaptured := false
        ; FIX: Re-apply theme after clearing value
        ApplyControlTheming(ShortcutEdit, "Edit", colors, isDark)
    }
    
    UpdateShortcutDisplay() {
        modifiers := []
        regularKeys := []
        
        ; Categorize pressed keys
        for key, _ in shortcutPressedKeys {
            if (key = "LCtrl" || key = "RCtrl")
                AddUniqueToArray(modifiers, "Ctrl")
            else if (key = "LAlt" || key = "RAlt") 
                AddUniqueToArray(modifiers, "Alt")
            else if (key = "LShift" || key = "RShift")
                AddUniqueToArray(modifiers, "Shift")
            else if (key = "LWin" || key = "RWin")
                AddUniqueToArray(modifiers, "Win")
            else
                AddUniqueToArray(regularKeys, key)
        }
        
        ; Build shortcut string
        shortcutText := ""
        if (modifiers.Length > 0)
            shortcutText := Join(modifiers, "+")
            
        if (regularKeys.Length > 0) {
            if (shortcutText != "")
                shortcutText .= "+"
            shortcutText .= Join(regularKeys, "+")
        }
        
        if (shortcutText != "")
            ShortcutEdit.Value := shortcutText
            ; FIX: Re-apply theme after updating value
            ApplyControlTheming(ShortcutEdit, "Edit", colors, isDark)
    }
    
    ShortcutKeyDown(key, *) {
        ; Only capture when ShortcutEdit has focus
        try {
            if (ControlGetFocus("ahk_id " AddAppGui.Hwnd) != ShortcutEdit.Name) {
                return
            }
        } catch {
            return
        }
        
        shortcutPressedKeys[key] := true
        UpdateShortcutDisplay()
    }
    
    ShortcutKeyUp(key, *) {
        if (shortcutPressedKeys.Has(key)) {
            shortcutPressedKeys.Delete(key)
        }
        
        ; If all keys released, finalize the shortcut
        if (shortcutPressedKeys.Count = 0 && ShortcutEdit.Value != "") {
            shortcutHasCaptured := true
        }
    }
    
    RegisterShortcutKey(key) {
        try {
            Hotkey("~*" key, ShortcutKeyDown.Bind(key))
            Hotkey("~*" key " up", ShortcutKeyUp.Bind(key))
        } catch {
            ; Key might already be registered or unavailable
        }
    }
    
    EnableShortcutCapture(*) {
        if (shortcutHotkeysRegistered) {
            return
        }
        
        ; Set up hotkey capture when ShortcutEdit has focus
        HotIfWinActive("ahk_id " AddAppGui.Hwnd)
        
        ; Register key capture hotkeys
        try {
            ; Modifier keys
            RegisterShortcutKey("LCtrl")
            RegisterShortcutKey("RCtrl") 
            RegisterShortcutKey("LAlt")
            RegisterShortcutKey("RAlt")
            RegisterShortcutKey("LShift")
            RegisterShortcutKey("RShift")
            RegisterShortcutKey("LWin")
            RegisterShortcutKey("RWin")
            
            ; Alphabet keys
            alphabet := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            Loop Parse, alphabet
                RegisterShortcutKey(A_LoopField)
                
            ; Number keys
            Loop 10
                RegisterShortcutKey(A_Index-1)
                
            ; Function keys
            Loop 12
                RegisterShortcutKey("F" A_Index)
                
            ; Other common keys
            otherKeys := ["Tab", "Space", "Delete", "Home", "End", "PgUp", "PgDn", "Insert"]
            For key in otherKeys
                RegisterShortcutKey(key)
                
            shortcutHotkeysRegistered := true
        } catch {
            ; Some hotkeys might fail to register
        }
        
        HotIf()
    }
    
    HandleListViewSelection(*) {
        selectedRow := LV.GetNext()
        
        if (selectedRow > 0) {
            ; Load selected shortcut into edit fields
            CategoryEdit.Value := LV.GetText(selectedRow, 1)
            CommandEdit.Value := LV.GetText(selectedRow, 2)
            ShortcutEdit.Value := LV.GetText(selectedRow, 3)
            DescriptionEdit.Value := LV.GetText(selectedRow, 4)
            
            ; FIX: Re-apply the theme to all edit fields after changing their values
            ApplyControlTheming(CategoryEdit, "Edit", colors, isDark)
            ApplyControlTheming(CommandEdit, "Edit", colors, isDark)
            ApplyControlTheming(ShortcutEdit, "Edit", colors, isDark)
            ApplyControlTheming(DescriptionEdit, "Edit", colors, isDark)

            isEditing := true
            UpdateButton.Enabled := true
            EnableShortcutCapture()
        } else {
            ClearEditFields()
            isEditing := false
            UpdateButton.Enabled := false
            selectedRow := 0
        }
    }
    
    HandleListViewDoubleClick(*) {
        selectedRow := LV.GetNext()
        if (selectedRow > 0) {
            HandleListViewSelection()
            CommandEdit.Focus()
        }
    }
    
    AddNewShortcut(*) {
        ; Clear all edit fields
        CategoryEdit.Value := "General"
        CommandEdit.Value := ""
        ShortcutEdit.Value := ""
        DescriptionEdit.Value := ""
        
        ; Reset editing state
        selectedRow := 0
        isEditing := false
        UpdateButton.Enabled := false
        
        ; Focus on command field for new entry
        CommandEdit.Focus()
        
        ; Enable shortcut capture
        EnableShortcutCapture()
    }
    
    UpdateSelectedShortcut(*) {
        if (selectedRow <= 0 || selectedRow > LV.GetCount()) {
            MsgBox("Please select a shortcut to update.", "No Selection", 48)
            return
        }
        
        ; Validate required fields
        if (CategoryEdit.Value == "" || CommandEdit.Value == "" || ShortcutEdit.Value == "") {
            MsgBox("Please fill in Category, Command, and Shortcut fields.", "Missing Information", 48)
            return
        }
        
        ; Update the ListView row
        LV.Modify(selectedRow, , CategoryEdit.Value, CommandEdit.Value, ShortcutEdit.Value, DescriptionEdit.Value)
        
        ; Clear fields and reset state
        ClearEditFields()
        selectedRow := 0
        isEditing := false
        UpdateButton.Enabled := false
        
        MsgBox("Shortcut updated successfully.", "Update Complete", 64)
    }
    
    DeleteSelectedShortcut(*) {
        if (selectedRow <= 0 || selectedRow > LV.GetCount()) {
            MsgBox("Please select a shortcut to delete.", "No Selection", 48)
            return
        }
        
        ; Get shortcut info for confirmation
        command := LV.GetText(selectedRow, 2)
        shortcut := LV.GetText(selectedRow, 3)
        
        result := MsgBox("Are you sure you want to delete this shortcut?`n`nCommand: " command "`nShortcut: " shortcut, "Confirm Delete", 52)
        
        if (result = "Yes") {
            LV.Delete(selectedRow)
            ClearEditFields()
            selectedRow := 0
            isEditing := false
            UpdateButton.Enabled := false
            MsgBox("Shortcut deleted successfully.", "Delete Complete", 64)
        }
    }
    
    SaveAllShortcuts(*) {
        ; Get all shortcuts from ListView
        shortcuts := []
        
        Loop LV.GetCount() {
            shortcut := {
                category: LV.GetText(A_Index, 1),
                command_name: LV.GetText(A_Index, 2),
                shortcut_key: LV.GetText(A_Index, 3),
                description: LV.GetText(A_Index, 4)
            }
            shortcuts.Push(shortcut)
        }
        
        ; Update the global shortcuts database
        global shortcuts_db
        if (shortcuts.Length > 0) {
            shortcuts_db[processName] := shortcuts
        } else if (shortcuts_db.Has(processName)) {
            shortcuts_db.Delete(processName)
        }
        
        ; Save to persistent storage
        SaveShortcutsToFile()
        
        ; Try to save to SQLite database if available
        if (InitializeSQLite()) {
            try {
                SaveMemoryToDatabase()
            } catch {
                ; Continue if SQLite save fails
            }
        }
        
        ; Update main window display if this is the current app
        global lastDetectedHwnd
        if (lastDetectedHwnd != 0 && WinExist("ahk_id " lastDetectedHwnd)) {
            currentProcess := WinGetProcessName("ahk_id " lastDetectedHwnd)
            if (currentProcess = processName) {
                UpdateShortcutsInfo()
            }
        }
        
        MsgBox("All shortcuts saved successfully!", "Save Complete", 64)
        AddAppGui.Destroy()
    }
    
    ShowBulkAddShortcutsDialog(parentGui, targetListView) {
        ; COMPREHENSIVE THEME CHECK BEFORE DIALOG CREATION
        ForceThemeCheck()
        
        try {
            BulkGui := CreateManagedDialog("+Owner" parentGui.Hwnd, "Bulk Add Shortcuts")
        } catch {
            BulkGui := Gui("+Owner" parentGui.Hwnd, "Bulk Add Shortcuts")
            ApplyDarkModeToGUI(BulkGui, "Bulk Add Dialog")
        }
        BulkGui.SetFont("s10")
        colors := GetThemeColors()
        isDark := IsWindowsDarkMode()
        
        instructionText := BulkGui.Add("Text", "x10 y10 w480", "Enter shortcuts in CSV format (one per line):")
        ApplyControlTheming(instructionText, "Text", colors, isDark)
        
        formatText := BulkGui.Add("Text", "x10 y30 w480", "Format: Category,Command,Shortcut,Description")
        formatText.SetFont("italic s9")
        ApplyControlTheming(formatText, "Text", colors, isDark)
        
        exampleText := BulkGui.Add("Text", "x10 y50 w480", "Example: File Operations,Save,Ctrl+S,Save the current document")
        exampleText.SetFont("italic s9")
        ApplyControlTheming(exampleText, "Text", colors, isDark)
        
        BulkEdit := BulkGui.Add("Edit", "x10 y80 w480 h200 Multi VScroll")
        ApplyControlTheming(BulkEdit, "Edit", colors, isDark)
        
        ; Pre-fill with example
        exampleContent := "File Operations,New,Ctrl+N,Create a new document`n"
        exampleContent .= "File Operations,Open,Ctrl+O,Open an existing document`n"
        exampleContent .= "Edit,Copy,Ctrl+C,Copy selected text`n"
        exampleContent .= "Edit,Paste,Ctrl+V,Paste from clipboard"
        BulkEdit.Value := exampleContent
        
        AddBulkButton := BulkGui.Add("Button", "Default x150 y300 w100", "&Add All")
        ApplyControlTheming(AddBulkButton, "Button", colors, isDark)
        
        CancelBulkButton := BulkGui.Add("Button", "x260 y300 w100", "&Cancel")
        ApplyControlTheming(CancelBulkButton, "Button", colors, isDark)
        
        ProcessBulkAdd(*) {
            content := BulkEdit.Value
            if (content = "") {
                MsgBox("Please enter some shortcuts to add.", "No Content", 48)
                return
            }
            
            lines := StrSplit(content, "`n", "`r")
            addedCount := 0
            
            for line in lines {
                line := Trim(line)
                if (line = "")
                    continue
                    
                parts := StrSplit(line, ",")
                if (parts.Length >= 3) {
                    category := Trim(parts[1])
                    command := Trim(parts[2])
                    shortcut := Trim(parts[3])
                    description := parts.Length >= 4 ? Trim(parts[4]) : ""
                    
                    if (category != "" && command != "" && shortcut != "") {
                        targetListView.Add(, category, command, shortcut, description)
                        addedCount++
                    }
                }
            }
            
            if (addedCount > 0) {
                MsgBox("Added " addedCount " shortcuts successfully!", "Bulk Add Complete", 64)
                BulkGui.Destroy()
            } else {
                MsgBox("No valid shortcuts found. Please check the format.", "No Shortcuts Added", 48)
            }
        }
        
        AddBulkButton.OnEvent("Click", ProcessBulkAdd)
        CancelBulkButton.OnEvent("Click", (*) => BulkGui.Destroy())
        BulkGui.OnEvent("Close", (*) => BulkGui.Destroy())
        
        RegisterGuiForEscapeHandling(BulkGui, (*) => BulkGui.Destroy())
        
        BulkGui.Show("w500 h340")
        ForceWindowToForeground(BulkGui.Hwnd)
        
        ; Final theme consistency check
        ForceThemeCheck()
        RefreshGUITheme(BulkGui, "Bulk Add Dialog")
        
        ; Apply enhanced text control theming
        ForceTextControlsRetheming(BulkGui)
    }
    
    ; Assign all event handlers
    AddNewButton.OnEvent("Click", AddNewShortcut)
    UpdateButton.OnEvent("Click", UpdateSelectedShortcut)
    BulkAddButton.OnEvent("Click", (*) => ShowBulkAddShortcutsDialog(AddAppGui, LV))
    DeleteButton.OnEvent("Click", DeleteSelectedShortcut)
    SaveButton.OnEvent("Click", SaveAllShortcuts)
    CancelButton.OnEvent("Click", (*) => AddAppGui.Destroy())
    
    ; ListView event handlers
    LV.OnEvent("ItemSelect", HandleListViewSelection)
    LV.OnEvent("DoubleClick", HandleListViewDoubleClick)
    
    ; Shortcut field event handlers  
    ShortcutEdit.OnEvent("Focus", EnableShortcutCapture)
    ClearShortcutBtn.OnEvent("Click", ClearShortcutField)
    
    RegisterGuiForEscapeHandling(AddAppGui, (*) => AddAppGui.Destroy())
    
    ; Show with comfortable initial size (larger than minimum for better usability)
    AddAppGui.Show("w620 h550")
    ForceWindowToForeground(AddAppGui.Hwnd)
    
    ; Final theme consistency check after showing
    ForceThemeCheck()
    RefreshGUITheme(AddAppGui, "Shortcut Management Dialog")
    
    ; Apply enhanced text control theming
    ForceTextControlsRetheming(AddAppGui)
}

;==============================================================================
; ENHANCED EXPORT FUNCTIONALITY SYSTEM WITH COMPREHENSIVE THEME CHECKING
;==============================================================================

/**
 * Display comprehensive export dialog with format and scope options and enhanced theme checking
 * @return {Void}
 */
ShowExportDialog(*) {
    global shortcuts_db, INSTALL_DIR, MyGui
    
    ; COMPREHENSIVE THEME CHECK BEFORE DIALOG CREATION
    ForceThemeCheck()
    
    try {
        ExportGui := CreateManagedDialog("+Owner" MyGui.Hwnd, "Export Shortcuts")
    } catch {
        ExportGui := Gui("+Owner" MyGui.Hwnd, "Export Shortcuts")
        ApplyDarkModeToGUI(ExportGui, "Export Dialog")
    }
    ExportGui.SetFont("s10")
    colors := GetThemeColors()
    isDark := IsWindowsDarkMode()
    
    titleText := ExportGui.Add("Text", "x10 y10 w380", "Export shortcuts database to file")
    ApplyControlTheming(titleText, "Text", colors, isDark)
    
    ; Export format selection
    formatLabel := ExportGui.Add("Text", "x10 y40 w100", "Export Format:")
    ApplyControlTheming(formatLabel, "Text", colors, isDark)
    
    FormatGroup := ExportGui.Add("GroupBox", "x10 y60 w380 h80", "File Format")
    ApplyControlTheming(FormatGroup, "GroupBox", colors, isDark)
    
    CSVRadio := ExportGui.Add("Radio", "x20 y85 w100 Checked", "&CSV Format")
    ApplyControlTheming(CSVRadio, "Button", colors, isDark)
    
    HTMLRadio := ExportGui.Add("Radio", "x130 y85 w100", "&HTML Format")
    ApplyControlTheming(HTMLRadio, "Button", colors, isDark)
    
    TextRadio := ExportGui.Add("Radio", "x240 y85 w100", "&Text Format")
    ApplyControlTheming(TextRadio, "Button", colors, isDark)
    
    ; Export scope selection
    ScopeGroup := ExportGui.Add("GroupBox", "x10 y150 w380 h80", "Export Scope")
    ApplyControlTheming(ScopeGroup, "GroupBox", colors, isDark)
    
    AllAppsRadio := ExportGui.Add("Radio", "x20 y175 w150 Checked", "&All Applications")
    ApplyControlTheming(AllAppsRadio, "Button", colors, isDark)
    
    CurrentAppRadio := ExportGui.Add("Radio", "x180 y175 w150", "&Current Application Only")
    ApplyControlTheming(CurrentAppRadio, "Button", colors, isDark)
    
    ; Enable current app option only if an application is captured
    global lastDetectedHwnd
    if (lastDetectedHwnd = 0 || !WinExist("ahk_id " lastDetectedHwnd)) {
        CurrentAppRadio.Enabled := false
        CurrentAppRadio.Text := "Current Application Only (none captured)"
    }
    
    ; Export configuration options
    OptionsGroup := ExportGui.Add("GroupBox", "x10 y240 w380 h60", "Options")
    ApplyControlTheming(OptionsGroup, "GroupBox", colors, isDark)
    
    IncludeDescCheck := ExportGui.Add("Checkbox", "x20 y265 w150 Checked", "Include &Descriptions")
    ApplyControlTheming(IncludeDescCheck, "Button", colors, isDark)
    
    GroupByCategoryCheck := ExportGui.Add("Checkbox", "x180 y265 w150 Checked", "&Group by Category")
    ApplyControlTheming(GroupByCategoryCheck, "Button", colors, isDark)
    
    ; Action buttons
    ExportNowButton := ExportGui.Add("Button", "Default x100 y320 w100", "&Export")
    ApplyControlTheming(ExportNowButton, "Button", colors, isDark)
    
    CancelExportButton := ExportGui.Add("Button", "x210 y320 w100", "&Cancel")
    ApplyControlTheming(CancelExportButton, "Button", colors, isDark)
    
    ExportNowButton.OnEvent("Click", PerformExport)
    CancelExportButton.OnEvent("Click", (*) => ExportGui.Destroy())
    ExportGui.OnEvent("Close", (*) => ExportGui.Destroy())
    
    RegisterGuiForEscapeHandling(ExportGui, (*) => ExportGui.Destroy())
    
    PerformExport(*) {
        ; Determine export format from user selection
        exportFormat := ""
        if (CSVRadio.Value)
            exportFormat := "CSV"
        else if (HTMLRadio.Value)
            exportFormat := "HTML"
        else if (TextRadio.Value)
            exportFormat := "TEXT"
        
        ; Determine export scope and options
        exportAll := AllAppsRadio.Value
        includeDesc := IncludeDescCheck.Value
        groupByCategory := GroupByCategoryCheck.Value
        
        ; Prepare export data based on scope
        if (exportAll) {
            exportData := shortcuts_db.Clone()
        } else {
            if (lastDetectedHwnd = 0 || !WinExist("ahk_id " lastDetectedHwnd)) {
                MsgBox("No application is currently captured.", "Export Error", 48)
                return
            }
            
            processName := WinGetProcessName("ahk_id " lastDetectedHwnd)
            exportData := Map()
            if (shortcuts_db.Has(processName))
                exportData[processName] := shortcuts_db[processName]
        }
        
        if (exportData.Count = 0) {
            MsgBox("No shortcuts found to export.", "Export Error", 48)
            return
        }
        
        ; Generate appropriate filename with timestamp
        timestamp := FormatTime(, "yyyyMMdd_HHmmss")
        scope := exportAll ? "all_shortcuts" : "current_app_shortcuts"
        
        switch exportFormat {
            case "CSV":
                filename := "shortcuts_export_" scope "_" timestamp ".csv"
                filter := "CSV Files (*.csv)"
            case "HTML":
                filename := "shortcuts_export_" scope "_" timestamp ".html"
                filter := "HTML Files (*.html)"
            case "TEXT":
                filename := "shortcuts_export_" scope "_" timestamp ".txt"
                filter := "Text Files (*.txt)"
        }
        
        ; Present file save dialog to user
        defaultPath := INSTALL_DIR "\Exports\" filename
        selectedFile := FileSelect("S16", defaultPath, "Save Export File", filter)
        
        if (selectedFile = "")
            return
        
        ; Execute export operation
        success := false
        try {
            switch exportFormat {
                case "CSV":
                    success := ExportToCSV(exportData, selectedFile, includeDesc, groupByCategory)
                case "HTML": 
                    success := ExportToHTML(exportData, selectedFile, includeDesc, groupByCategory)
                case "TEXT":
                    success := ExportToText(exportData, selectedFile, includeDesc, groupByCategory)
            }
            
            if (success) {
                MsgBox("Export completed successfully!`n`nFile: " selectedFile, "Export Complete", 64)
                ExportGui.Destroy()
            } else {
                MsgBox("Export failed. Please check file permissions and try again.", "Export Error", 16)
            }
        } catch Error as e {
            MsgBox("Export error: " e.Message, "Export Error", 16)
        }
    }
    
    ExportGui.Show("w400 h360")
    ForceWindowToForeground(ExportGui.Hwnd)
    
    ; Final theme consistency check after showing
    ForceThemeCheck()
    RefreshGUITheme(ExportGui, "Export Dialog")
    
    ; Apply enhanced text control theming
    ForceTextControlsRetheming(ExportGui)
}

/**
 * Export shortcuts to CSV format with proper escaping
 * @param {Map} exportData - Shortcuts data to export
 * @param {String} filePath - Output file path
 * @param {Boolean} includeDesc - Include description column
 * @param {Boolean} groupByCategory - Group entries by category
 * @return {Boolean} True if export successful
 */
ExportToCSV(exportData, filePath, includeDesc := true, groupByCategory := true) {
    try {
        if FileExist(filePath)
            FileDelete(filePath)
        
        ; Create CSV header row
        header := "Application,Category,Command,Shortcut"
        if (includeDesc)
            header .= ",Description"
        header .= "`n"
        
        FileAppend(header, filePath)
        
        ; Process each application's shortcuts
        for processName, shortcuts in exportData {
            appDisplayName := GetAppDisplayName(processName)
            
            if (groupByCategory) {
                ; Organize shortcuts by category
                categorized := Map()
                for shortcut in shortcuts {
                    category := shortcut.HasOwnProp("category") ? shortcut.category : "General"
                    if (!categorized.Has(category))
                        categorized[category] := []
                    categorized[category].Push(shortcut)
                }
                
                ; Export each category group
                for category, categoryShortcuts in categorized {
                    for shortcut in categoryShortcuts {
                        line := EscapeCSV(appDisplayName) "," EscapeCSV(category) "," 
                        line .= EscapeCSV(shortcut.HasOwnProp("command_name") ? shortcut.command_name : "") ","
                        line .= EscapeCSV(shortcut.HasOwnProp("shortcut_key") ? shortcut.shortcut_key : "")
                        
                        if (includeDesc)
                            line .= "," EscapeCSV(shortcut.HasOwnProp("description") ? shortcut.description : "")
                        
                        line .= "`n"
                        FileAppend(line, filePath)
                    }
                }
            } else {
                ; Export shortcuts in original order
                for shortcut in shortcuts {
                    category := shortcut.HasOwnProp("category") ? shortcut.category : "General"
                    line := EscapeCSV(appDisplayName) "," EscapeCSV(category) ","
                    line .= EscapeCSV(shortcut.HasOwnProp("command_name") ? shortcut.command_name : "") ","
                    line .= EscapeCSV(shortcut.HasOwnProp("shortcut_key") ? shortcut.shortcut_key : "")
                    
                    if (includeDesc)
                        line .= "," EscapeCSV(shortcut.HasOwnProp("description") ? shortcut.description : "")
                    
                    line .= "`n"
                    FileAppend(line, filePath)
                }
            }
        }
        
        return true
    } catch {
        return false
    }
}

/**
 * Export shortcuts to HTML format with professional styling
 * @param {Map} exportData - Shortcuts data to export
 * @param {String} filePath - Output file path
 * @param {Boolean} includeDesc - Include description column
 * @param {Boolean} groupByCategory - Group entries by category
 * @return {Boolean} True if export successful
 */
ExportToHTML(exportData, filePath, includeDesc := true, groupByCategory := true) {
    try {
        if FileExist(filePath)
            FileDelete(filePath)
        
        ; Build complete HTML document
        html := "<!DOCTYPE html>`n<html>`n<head>`n"
        html .= "<meta charset='UTF-8'>`n"
        html .= "<title>Shortcuts Export - " FormatTime(, "yyyy-MM-dd HH:mm:ss") "</title>`n"
        html .= "<style>`n"
        html .= "body { font-family: Arial, sans-serif; margin: 20px; }`n"
        html .= "h1 { color: #333; }`n"
        html .= "h2 { color: #666; border-bottom: 2px solid #ccc; }`n"
        html .= "h3 { color: #888; }`n"
        html .= "table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }`n"
        html .= "th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }`n"
        html .= "th { background-color: #f2f2f2; font-weight: bold; }`n"
        html .= "tr:nth-child(even) { background-color: #f9f9f9; }`n"
        html .= ".shortcut { font-family: 'Courier New', monospace; font-weight: bold; }`n"
        html .= "</style>`n</head>`n<body>`n"
        
        ; Document header
        html .= "<h1>Application Shortcuts Export</h1>`n"
        html .= "<p>Generated on: " FormatTime(, "yyyy-MM-dd HH:mm:ss") "</p>`n"
        
        ; Process each application
        for processName, shortcuts in exportData {
            appDisplayName := GetAppDisplayName(processName)
            html .= "<h2>" EscapeHTML(appDisplayName) "</h2>`n"
            
            if (groupByCategory) {
                ; Group and display by category
                categorized := Map()
                for shortcut in shortcuts {
                    category := shortcut.HasOwnProp("category") ? shortcut.category : "General"
                    if (!categorized.Has(category))
                        categorized[category] := []
                    categorized[category].Push(shortcut)
                }
                
                for category, categoryShortcuts in categorized {
                    html .= "<h3>" EscapeHTML(category) "</h3>`n"
                    html .= "<table>`n<tr><th>Command</th><th>Shortcut</th>"
                    if (includeDesc)
                        html .= "<th>Description</th>"
                    html .= "</tr>`n"
                    
                    for shortcut in categoryShortcuts {
                        html .= "<tr><td>" EscapeHTML(shortcut.HasOwnProp("command_name") ? shortcut.command_name : "") "</td>"
                        html .= "<td class='shortcut'>" EscapeHTML(shortcut.HasOwnProp("shortcut_key") ? shortcut.shortcut_key : "") "</td>"
                        if (includeDesc)
                            html .= "<td>" EscapeHTML(shortcut.HasOwnProp("description") ? shortcut.description : "") "</td>"
                        html .= "</tr>`n"
                    }
                    html .= "</table>`n"
                }
            } else {
                ; Display in single table
                html .= "<table>`n<tr><th>Category</th><th>Command</th><th>Shortcut</th>"
                if (includeDesc)
                    html .= "<th>Description</th>"
                html .= "</tr>`n"
                
                for shortcut in shortcuts {
                    category := shortcut.HasOwnProp("category") ? shortcut.category : "General"
                    html .= "<tr><td>" EscapeHTML(category) "</td>"
                    html .= "<td>" EscapeHTML(shortcut.HasOwnProp("command_name") ? shortcut.command_name : "") "</td>"
                    html .= "<td class='shortcut'>" EscapeHTML(shortcut.HasOwnProp("shortcut_key") ? shortcut.shortcut_key : "") "</td>"
                    if (includeDesc)
                        html .= "<td>" EscapeHTML(shortcut.HasOwnProp("description") ? shortcut.description : "") "</td>"
                    html .= "</tr>`n"
                }
                html .= "</table>`n"
            }
        }
        
        html .= "</body>`n</html>"
        
        FileAppend(html, filePath)
        return true
    } catch {
        return false
    }
}

/**
 * Export shortcuts to plain text format
 * @param {Map} exportData - Shortcuts data to export
 * @param {String} filePath - Output file path
 * @param {Boolean} includeDesc - Include descriptions
 * @param {Boolean} groupByCategory - Group entries by category
 * @return {Boolean} True if export successful
 */
ExportToText(exportData, filePath, includeDesc := true, groupByCategory := true) {
    try {
        if FileExist(filePath)
            FileDelete(filePath)
        
        ; Build text document
        content := "Application Shortcuts Export`n"
        content .= "Generated on: " FormatTime(, "yyyy-MM-dd HH:mm:ss") "`n"
        content .= "================================================================`n`n"
        
        for processName, shortcuts in exportData {
            appDisplayName := GetAppDisplayName(processName)
            content .= appDisplayName "`n"
            content .= StrReplace(appDisplayName, "", "=", , -1) "`n`n"
            
            if (groupByCategory) {
                ; Group by category
                categorized := Map()
                for shortcut in shortcuts {
                    category := shortcut.HasOwnProp("category") ? shortcut.category : "General"
                    if (!categorized.Has(category))
                        categorized[category] := []
                    categorized[category].Push(shortcut)
                }
                
                for category, categoryShortcuts in categorized {
                    content .= category "`n"
                    content .= StrReplace(category, "", "-", , -1) "`n"
                    
                    for shortcut in categoryShortcuts {
                        command := shortcut.HasOwnProp("command_name") ? shortcut.command_name : ""
                        keys := shortcut.HasOwnProp("shortcut_key") ? shortcut.shortcut_key : ""
                        desc := shortcut.HasOwnProp("description") ? shortcut.description : ""
                        
                        content .= command "`n"
                        content .= "  Shortcut: " keys "`n"
                        if (includeDesc && desc != "")
                            content .= "  Description: " desc "`n"
                        content .= "`n"
                    }
                    content .= "`n"
                }
            } else {
                for shortcut in shortcuts {
                    category := shortcut.HasOwnProp("category") ? shortcut.category : "General"
                    command := shortcut.HasOwnProp("command_name") ? shortcut.command_name : ""
                    keys := shortcut.HasOwnProp("shortcut_key") ? shortcut.shortcut_key : ""
                    desc := shortcut.HasOwnProp("description") ? shortcut.description : ""
                    
                    content .= command " [" category "]`n"
                    content .= "  Shortcut: " keys "`n"
                    if (includeDesc && desc != "")
                        content .= "  Description: " desc "`n"
                    content .= "`n"
                }
            }
            content .= "`n"
        }
        
        FileAppend(content, filePath)
        return true
    } catch {
        return false
    }
}

/**
 * Escape text for safe CSV output
 * @param {String} text - Text to escape for CSV format
 * @return {String} Properly escaped CSV text
 */
EscapeCSV(text) {
    if (text = "")
        return '""'
    
    ; Escape quotes and wrap in quotes if necessary
    if (InStr(text, '"') || InStr(text, ",") || InStr(text, "`n") || InStr(text, "`r")) {
        text := StrReplace(text, '"', '""')
        return '"' text '"'
    }
    return text
}

/**
 * Escape text for safe HTML output
 * @param {String} text - Text to escape for HTML format
 * @return {String} Properly escaped HTML text
 */
EscapeHTML(text) {
    if (text = "")
        return ""
    
    ; Escape HTML special characters
    text := StrReplace(text, "&", "&amp;")
    text := StrReplace(text, "<", "&lt;")
    text := StrReplace(text, ">", "&gt;")
    text := StrReplace(text, '"', "&quot;")
    text := StrReplace(text, "'", "&#39;")
    return text
}
;==============================================================================
; INITIALIZATION AND STARTUP WITH ENHANCED THEME SUPPORT
;==============================================================================

/**
 * Global message handler for WM_CTLCOLORSTATIC (0x0138).
 * This function is called whenever a static or readonly edit control
 * is about to be drawn.
 */

; Register Windows message handler for Escape key processing
OnMessage(0x0100, WM_KEYDOWN)

; Initialize enhanced stable responsive dark mode and theme monitoring system
EnableDarkModeForApp()
CurrentTheme := IsWindowsDarkMode()

; Set up enhanced theme monitoring BEFORE other initialization
SetupEnhancedThemeMonitoring()

; Execute installation system initialization
InitializeInstallation()

;Initialize database system with embedded file support
InitDatabase()

; Create and initialize main application GUI with enhanced responsive theme support
InitializeGUIWithThemeSupport()

;Initialize system tray icon and menu
InitializeSystemTray()

; Register global hotkeys dynamically
RegisterGlobalHotkeys()

; Final theme refresh on startup to ensure everything is properly themed
ScheduleFullThemeRefresh()