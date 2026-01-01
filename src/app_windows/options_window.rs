use windows::core::{PCWSTR, w};
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM, LRESULT, HINSTANCE};
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetWindowLongPtrW, RegisterClassW,
    KillTimer, PostMessageW, SendMessageW, SetTimer, SetWindowLongPtrW, SetForegroundWindow, SetWindowTextW,
    GWLP_USERDATA, WM_CREATE, WM_DESTROY, WM_NCDESTROY, WM_CLOSE, WM_COMMAND,
    WM_KEYDOWN, WM_SETFONT, WM_APP, WM_NEXTDLGCTL, WM_TIMER, WM_SETFOCUS,
    WS_CAPTION, WS_SYSMENU, WS_VISIBLE, WS_CHILD, WS_TABSTOP, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME,
    CW_USEDEFAULT, HMENU, WNDCLASSW, WS_EX_CLIENTEDGE,
    BS_DEFPUSHBUTTON, BS_AUTOCHECKBOX, BM_SETCHECK, BM_GETCHECK, BM_CLICK,
    CB_ADDSTRING, CB_RESETCONTENT, CB_SETCURSEL, CB_GETCURSEL, CB_GETITEMDATA, CB_SETITEMDATA, CBS_DROPDOWNLIST,
    CB_GETDROPPEDSTATE, GetParent, MSG, GetWindowTextLengthW, GetWindowTextW, ShowWindow,
    CREATESTRUCTW, LoadCursorW, IDC_ARROW, WINDOW_STYLE, CBN_SELCHANGE, SW_HIDE, SW_SHOW,
    ES_AUTOHSCROLL
};
use windows::Win32::UI::Controls::{WC_BUTTON, WC_STATIC, WC_COMBOBOXW, BST_CHECKED};
use windows::Win32::Graphics::Gdi::{HBRUSH, COLOR_WINDOW, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::{GetFocus, SetFocus, VK_RETURN, VK_ESCAPE, EnableWindow};
use crate::{with_state, rebuild_menus, refresh_voice_panel};
use crate::editor_manager::apply_word_wrap_to_all_edits;
use crate::settings::{Language, OpenBehavior, VoiceInfo, TtsEngine, save_settings, TRUSTED_CLIENT_TOKEN, VOICE_LIST_URL};
use crate::accessibility::{to_wide, handle_accessibility};
use std::thread;

const OPTIONS_CLASS_NAME: &str = "NovapadOptions";
const OPTIONS_ID_LANG: usize = 6001;
const OPTIONS_ID_OPEN: usize = 6002;
const OPTIONS_ID_TTS_ENGINE: usize = 6012;
const OPTIONS_ID_VOICE: usize = 6003;
const OPTIONS_ID_MULTILINGUAL: usize = 6004;
const OPTIONS_ID_TTS_TUNING: usize = 6014;
const OPTIONS_ID_SPLIT_ON_NEWLINE: usize = 6007;
const OPTIONS_ID_WORD_WRAP: usize = 6008;
const OPTIONS_ID_MOVE_CURSOR: usize = 6009;
const OPTIONS_ID_AUDIO_SKIP: usize = 6010;
const OPTIONS_ID_AUDIO_SPLIT: usize = 6011;
const OPTIONS_ID_AUDIO_SPLIT_TEXT: usize = 6013;
const OPTIONS_ID_AUDIO_SPLIT_REQUIRE_NEWLINE: usize = 6016;
const OPTIONS_ID_CHECK_UPDATES: usize = 6015;
const OPTIONS_ID_OK: usize = 6005;
const OPTIONS_ID_CANCEL: usize = 6006;
const OPTIONS_FOCUS_LANG_MSG: u32 = WM_APP + 30;
const OPTIONS_FOCUS_LANG_TIMER_ID: usize = 1;

const WM_TTS_VOICES_LOADED: u32 = WM_APP + 2;
const WM_TTS_SAPI_VOICES_LOADED: u32 = WM_APP + 8;
const AUDIOBOOK_SPLIT_BY_TEXT: u32 = u32::MAX;

pub unsafe fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
        let focus = GetFocus();
        if GetParent(focus) == hwnd {
             let dropped = SendMessageW(focus, CB_GETDROPPEDSTATE, WPARAM(0), LPARAM(0)).0 != 0;
             if !dropped {
                 let _ = with_options_state(hwnd, |state| {
                     let _ = SendMessageW(hwnd, WM_COMMAND, WPARAM(OPTIONS_ID_OK | (0 << 16)), LPARAM(state.ok_button.0));
                 });
                 return true;
             }
        }
    }
    handle_accessibility(hwnd, msg)
}

struct OptionsDialogState {
    parent: HWND,
    label_language: HWND,
    combo_lang: HWND,
    combo_open: HWND,
    combo_tts_engine: HWND,
    combo_voice: HWND,
    combo_audio_skip: HWND,
    combo_audio_split: HWND,
    label_audio_split_text: HWND,
    edit_audio_split_text: HWND,
    checkbox_audio_split_requires_newline: HWND,
    checkbox_multilingual: HWND,
    button_tts_tuning: HWND,
    checkbox_split_on_newline: HWND,
    checkbox_word_wrap: HWND,
    checkbox_move_cursor: HWND,
    checkbox_check_updates: HWND,
    ok_button: HWND,
}

struct OptionsLabels {
    title: &'static str,
    label_language: &'static str,
    label_open: &'static str,
    label_tts_engine: &'static str,
    label_voice: &'static str,
    label_multilingual: &'static str,
    label_tts_tuning: &'static str,
    label_split_on_newline: &'static str,
    label_word_wrap: &'static str,
    label_move_cursor: &'static str,
    label_check_updates: &'static str,
    label_audio_skip: &'static str,
    label_audio_split: &'static str,
    label_audio_split_text: &'static str,
    label_audio_split_requires_newline: &'static str,
    lang_it: &'static str,
    lang_en: &'static str,
    open_new_tab: &'static str,
    open_new_window: &'static str,
    engine_edge: &'static str,
    engine_sapi5: &'static str,
    split_none: &'static str,
    split_by_text: &'static str,
    split_parts: &'static str,
    ok: &'static str,
    cancel: &'static str,
    voices_empty: &'static str,
}

fn options_labels(language: Language) -> OptionsLabels {
    match language {
        Language::Italian => OptionsLabels {
            title: "Opzioni",
            label_language: "Lingua interfaccia:",
            label_open: "Apertura file:",
            label_tts_engine: "Sistema sintesi vocale:",
            label_voice: "Voce:",
            label_multilingual: "Mostra solo voci multilingua",
            label_tts_tuning: "Scegli tono, velocita' e volume",
            label_split_on_newline: "Spezza la lettura quando si va a capo",
            label_word_wrap: "A capo automatico nella finestra",
            label_move_cursor: "Sposta il cursore durante la lettura",
            label_check_updates: "Controlla aggiornamenti all'avvio",
            label_audio_skip: "Spostamento MP3 (frecce):",
            label_audio_split: "Dividi l'audiolibro in:",
            label_audio_split_text: "Testo per divisione:",
            label_audio_split_requires_newline: "Il testo deve iniziare a capo",
            lang_it: "Italiano",
            lang_en: "Inglese",
            open_new_tab: "Apri file in nuovo tab",
            open_new_window: "Apri file in nuova finestra",
            engine_edge: "Voci Microsoft",
            engine_sapi5: "SAPI5",
            split_none: "Nessuna divisione",
            split_by_text: "In base al testo",
            split_parts: "parti",
            ok: "OK",
            cancel: "Annulla",
            voices_empty: "Nessuna voce disponibile",
        },
        Language::English => OptionsLabels {
            title: "Options",
            label_language: "Interface language:",
            label_open: "Open behavior:",
            label_tts_engine: "Text-to-Speech System:",
            label_voice: "Voice:",
            label_multilingual: "Show only multilingual voices",
            label_tts_tuning: "Choose pitch, speed, and volume",
            label_split_on_newline: "Split reading on new lines",
            label_word_wrap: "Word wrap in editor",
            label_move_cursor: "Move cursor during reading",
            label_check_updates: "Check updates on startup",
            label_audio_skip: "MP3 skip interval:",
            label_audio_split: "Split audiobook into:",
            label_audio_split_text: "Split marker text:",
            label_audio_split_requires_newline: "Require the marker at line start",
            lang_it: "Italian",
            lang_en: "English",
            open_new_tab: "Open files in new tab",
            open_new_window: "Open files in new window",
            engine_edge: "Microsoft Voices",
            engine_sapi5: "SAPI5",
            split_none: "No split",
            split_by_text: "Based on text",
            split_parts: "parts",
            ok: "OK",
            cancel: "Cancel",
            voices_empty: "No voices available",
        },
    }
}

pub unsafe fn open(parent: HWND) {
    let existing = with_state(parent, |state| state.options_dialog).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }

    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(OPTIONS_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(LoadCursorW(None, IDC_ARROW).unwrap_or_default().0),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(options_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let labels = options_labels(language);
    let title = to_wide(labels.title);

    let dialog = CreateWindowExW(
        WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(title.as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        520,
        560, // Increased height
        parent,
        None,
        hinstance,
        Some(parent.0 as *const std::ffi::c_void),
    );

    if dialog.0 != 0 {
        let _ = with_state(parent, |state| {
            state.options_dialog = dialog;
        });
        EnableWindow(parent, true);
        SetForegroundWindow(dialog);
        ensure_voice_lists_loaded(parent, language);
    }
}

pub unsafe fn refresh_voices(hwnd: HWND) {
    let (parent, combo_voice, combo_engine, checkbox) = match with_options_state(hwnd, |state| {
        (state.parent, state.combo_voice, state.combo_tts_engine, state.checkbox_multilingual)
    }) {
        Some(values) => values,
        None => return,
    };
    let settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();
    
    // Determine current engine from combo if possible, otherwise settings
    let engine_sel = SendMessageW(combo_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let engine = if engine_sel >= 0 {
        if engine_sel == 1 { TtsEngine::Sapi5 } else { TtsEngine::Edge }
    } else {
        settings.tts_engine
    };

    let voices = with_state(parent, |state| {
        match engine {
            TtsEngine::Edge => state.edge_voices.clone(),
            TtsEngine::Sapi5 => state.sapi_voices.clone(),
        }
    }).unwrap_or_default();

    let labels = options_labels(settings.language);
    let only_multilingual = SendMessageW(checkbox, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32
        == BST_CHECKED.0;
    
    // Multilingual checkbox only relevant for Edge voices?
    // SAPI voices usually don't have "Multilingual" in name in the same way, but let's keep logic if applicable.
    // Generally assume SAPI voices are local and we list all.
    let filter_multilingual = if engine == TtsEngine::Edge { only_multilingual } else { false };
    
    // Disable multilingual checkbox for SAPI
    EnableWindow(checkbox, engine == TtsEngine::Edge);

    // If switching engine, we might not have the correct "selected" voice in settings yet if we haven't saved.
    // But we pass settings.tts_voice. If it's an ID from other engine, it won't match, so it selects default/first.
    populate_voice_combo(combo_voice, &voices, &settings.tts_voice, filter_multilingual, &labels);
}

unsafe extern "system" fn options_wndproc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let parent = HWND((*create_struct).lpCreateParams as isize);
            let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
            let labels = options_labels(language);

            let hfont = with_state(parent, |state| state.hfont).unwrap_or(HFONT(0));
            
            let mut y = 20;
            let label_lang = CreateWindowExW(Default::default(), WC_STATIC, PCWSTR(to_wide(labels.label_language).as_ptr()), WS_CHILD | WS_VISIBLE, 20, y, 140, 20, hwnd, HMENU(0), HINSTANCE(0), None);
            let combo_lang = CreateWindowExW(WS_EX_CLIENTEDGE, WC_COMBOBOXW, PCWSTR::null(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32), 170, y - 2, 300, 120, hwnd, HMENU(OPTIONS_ID_LANG as isize), HINSTANCE(0), None);
            y += 40;

            let label_open = CreateWindowExW(Default::default(), WC_STATIC, PCWSTR(to_wide(labels.label_open).as_ptr()), WS_CHILD | WS_VISIBLE, 20, y, 140, 20, hwnd, HMENU(0), HINSTANCE(0), None);
            let combo_open = CreateWindowExW(WS_EX_CLIENTEDGE, WC_COMBOBOXW, PCWSTR::null(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32), 170, y - 2, 300, 120, hwnd, HMENU(OPTIONS_ID_OPEN as isize), HINSTANCE(0), None);
            y += 40;

            let label_tts_engine = CreateWindowExW(Default::default(), WC_STATIC, PCWSTR(to_wide(labels.label_tts_engine).as_ptr()), WS_CHILD | WS_VISIBLE, 20, y, 140, 20, hwnd, HMENU(0), HINSTANCE(0), None);
            let combo_tts_engine = CreateWindowExW(WS_EX_CLIENTEDGE, WC_COMBOBOXW, PCWSTR::null(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32), 170, y - 2, 300, 120, hwnd, HMENU(OPTIONS_ID_TTS_ENGINE as isize), HINSTANCE(0), None);
            y += 40;

            let label_voice = CreateWindowExW(Default::default(), WC_STATIC, PCWSTR(to_wide(labels.label_voice).as_ptr()), WS_CHILD | WS_VISIBLE, 20, y, 140, 20, hwnd, HMENU(0), HINSTANCE(0), None);
            let combo_voice = CreateWindowExW(WS_EX_CLIENTEDGE, WC_COMBOBOXW, PCWSTR::null(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32), 170, y - 2, 300, 140, hwnd, HMENU(OPTIONS_ID_VOICE as isize), HINSTANCE(0), None);
            y += 40;

            let checkbox_multilingual = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.label_multilingual).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32), 170, y, 300, 20, hwnd, HMENU(OPTIONS_ID_MULTILINGUAL as isize), HINSTANCE(0), None);
            y += 28;

            let button_tts_tuning = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.label_tts_tuning).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP, 170, y, 300, 26, hwnd, HMENU(OPTIONS_ID_TTS_TUNING as isize), HINSTANCE(0), None);
            y += 36;

            let label_audio_skip = CreateWindowExW(Default::default(), WC_STATIC, PCWSTR(to_wide(labels.label_audio_skip).as_ptr()), WS_CHILD | WS_VISIBLE, 20, y, 140, 20, hwnd, HMENU(0), HINSTANCE(0), None);
            let combo_audio_skip = CreateWindowExW(WS_EX_CLIENTEDGE, WC_COMBOBOXW, PCWSTR::null(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32), 170, y - 2, 300, 140, hwnd, HMENU(OPTIONS_ID_AUDIO_SKIP as isize), HINSTANCE(0), None);
            y += 40;

            let label_audio_split = CreateWindowExW(Default::default(), WC_STATIC, PCWSTR(to_wide(labels.label_audio_split).as_ptr()), WS_CHILD | WS_VISIBLE, 20, y, 140, 20, hwnd, HMENU(0), HINSTANCE(0), None);
            let combo_audio_split = CreateWindowExW(WS_EX_CLIENTEDGE, WC_COMBOBOXW, PCWSTR::null(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32), 170, y - 2, 300, 140, hwnd, HMENU(OPTIONS_ID_AUDIO_SPLIT as isize), HINSTANCE(0), None);
            y += 34;

            let label_audio_split_text = CreateWindowExW(Default::default(), WC_STATIC, PCWSTR(to_wide(labels.label_audio_split_text).as_ptr()), WS_CHILD | WS_VISIBLE, 20, y, 140, 20, hwnd, HMENU(0), HINSTANCE(0), None);
            let edit_audio_split_text = CreateWindowExW(WS_EX_CLIENTEDGE, w!("EDIT"), PCWSTR::null(), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(ES_AUTOHSCROLL as u32), 170, y - 2, 300, 22, hwnd, HMENU(OPTIONS_ID_AUDIO_SPLIT_TEXT as isize), HINSTANCE(0), None);
            y += 34;

            let checkbox_audio_split_requires_newline = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.label_audio_split_requires_newline).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32), 170, y, 300, 20, hwnd, HMENU(OPTIONS_ID_AUDIO_SPLIT_REQUIRE_NEWLINE as isize), HINSTANCE(0), None);
            y += 24;

            let checkbox_split_on_newline = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.label_split_on_newline).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32), 170, y, 300, 20, hwnd, HMENU(OPTIONS_ID_SPLIT_ON_NEWLINE as isize), HINSTANCE(0), None);
            y += 24;

            let checkbox_word_wrap = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.label_word_wrap).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32), 170, y, 300, 20, hwnd, HMENU(OPTIONS_ID_WORD_WRAP as isize), HINSTANCE(0), None);
            y += 24;

            let checkbox_move_cursor = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.label_move_cursor).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32), 170, y, 300, 20, hwnd, HMENU(OPTIONS_ID_MOVE_CURSOR as isize), HINSTANCE(0), None);
            y += 24;

            let checkbox_check_updates = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.label_check_updates).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_AUTOCHECKBOX as u32), 170, y, 300, 20, hwnd, HMENU(OPTIONS_ID_CHECK_UPDATES as isize), HINSTANCE(0), None);
            y += 40;

            let ok_button = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.ok).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32), 280, y, 90, 28, hwnd, HMENU(OPTIONS_ID_OK as isize), HINSTANCE(0), None);
            let cancel_button = CreateWindowExW(Default::default(), WC_BUTTON, PCWSTR(to_wide(labels.cancel).as_ptr()), WS_CHILD | WS_VISIBLE | WS_TABSTOP, 380, y, 90, 28, hwnd, HMENU(OPTIONS_ID_CANCEL as isize), HINSTANCE(0), None);

            for control in [label_lang, combo_lang, label_open, combo_open, label_tts_engine, combo_tts_engine, label_voice, combo_voice, label_audio_skip, combo_audio_skip, label_audio_split, combo_audio_split, label_audio_split_text, edit_audio_split_text, checkbox_audio_split_requires_newline, checkbox_multilingual, button_tts_tuning, checkbox_split_on_newline, checkbox_word_wrap, checkbox_move_cursor, checkbox_check_updates, ok_button, cancel_button] {
                if control.0 != 0 && hfont.0 != 0 {
                    let _ = SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

            let dialog_state = Box::new(OptionsDialogState {
                parent,
                label_language: label_lang,
                combo_lang,
                combo_open,
                combo_tts_engine,
                combo_voice,
                combo_audio_skip,
                combo_audio_split,
                label_audio_split_text,
                edit_audio_split_text,
                checkbox_audio_split_requires_newline,
                checkbox_multilingual,
                button_tts_tuning,
                checkbox_split_on_newline,
                checkbox_word_wrap,
                checkbox_move_cursor,
                checkbox_check_updates,
                ok_button,
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(dialog_state) as isize);
            initialize_options_dialog(hwnd);
            SetFocus(combo_lang);
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xffff) as usize;
            let code = (wparam.0 >> 16) as u32;
            match cmd_id {
                OPTIONS_ID_OK => {
                    let focus = GetFocus();
                    let is_tuning = with_options_state(hwnd, |state| focus == state.button_tts_tuning).unwrap_or(false);
                    if is_tuning {
                        let parent = with_options_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                        if parent.0 != 0 {
                            crate::app_windows::tts_tuning_window::open(parent, hwnd);
                        }
                        LRESULT(0)
                    } else {
                        apply_options_dialog(hwnd);
                        LRESULT(0)
                    }
                }
                OPTIONS_ID_CANCEL | 2 => {
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                OPTIONS_ID_MULTILINGUAL => {
                    refresh_voices(hwnd);
                    LRESULT(0)
                }
                OPTIONS_ID_TTS_TUNING => {
                    let parent = with_options_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                    if parent.0 != 0 {
                        crate::app_windows::tts_tuning_window::open(parent, hwnd);
                    }
                    LRESULT(0)
                }
                OPTIONS_ID_TTS_ENGINE => {
                    if code == CBN_SELCHANGE {
                        // When engine changes, verify if we need to load SAPI voices
                        let combo = with_options_state(hwnd, |s| s.combo_tts_engine).unwrap_or(HWND(0));
                        let sel = SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
                        if sel == 1 { // SAPI
                            let parent = with_options_state(hwnd, |s| s.parent).unwrap_or(HWND(0));
                             let has_sapi = with_state(parent, |s| !s.sapi_voices.is_empty()).unwrap_or(false);
                             if !has_sapi {
                                 let lang = with_state(parent, |s| s.settings.language).unwrap_or_default();
                                 ensure_sapi_voices_loaded(parent, lang);
                             }
                        }
                        refresh_voices(hwnd);
                    }
                    LRESULT(0)
                }
                OPTIONS_ID_AUDIO_SPLIT => {
                    if code == CBN_SELCHANGE {
                        update_audio_split_text_visibility(hwnd);
                    }
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        OPTIONS_FOCUS_LANG_MSG => {
            focus_language_combo_once(hwnd);
            LRESULT(0)
        }
        WM_TIMER => {
            if wparam.0 == OPTIONS_FOCUS_LANG_TIMER_ID {
                let _ = KillTimer(hwnd, OPTIONS_FOCUS_LANG_TIMER_ID);
                focus_language_combo_once(hwnd);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == VK_RETURN.0 as u32 {
                let focus = GetFocus();
                let is_voice = with_options_state(hwnd, |state| focus == state.combo_voice).unwrap_or(false);
                if is_voice {
                    apply_options_dialog(hwnd);
                    return LRESULT(0);
                }
                let is_tuning = with_options_state(hwnd, |state| focus == state.button_tts_tuning).unwrap_or(false);
                if is_tuning {
                    let _ = SendMessageW(focus, BM_CLICK, WPARAM(0), LPARAM(0));
                    return LRESULT(0);
                }
            } else if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                let _ = DestroyWindow(hwnd);
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_DESTROY => {
            let parent = with_options_state(hwnd, |state| state.parent).unwrap_or(HWND(0));
            if parent.0 != 0 {
                EnableWindow(parent, true);
                SetForegroundWindow(parent);
                SetFocus(parent);
                if let Some(edit) = crate::get_active_edit(parent) {
                    SetFocus(edit);
                }
                let _ = PostMessageW(parent, crate::WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
                let _ = with_state(parent, |state| {
                    state.options_dialog = HWND(0);
                });
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut OptionsDialogState;
            if !ptr.is_null() {
                let _ = Box::from_raw(ptr);
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            let _ = DestroyWindow(hwnd);
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_options_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut OptionsDialogState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut OptionsDialogState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}

unsafe fn initialize_options_dialog(hwnd: HWND) {
    let (parent, combo_lang, combo_open, combo_tts_engine, _combo_voice, combo_audio_skip, combo_audio_split, _label_audio_split_text, edit_audio_split_text, checkbox_audio_split_requires_newline, checkbox_multilingual, _button_tts_tuning, checkbox_split_on_newline, checkbox_word_wrap, checkbox_move_cursor, checkbox_check_updates) = match with_options_state(hwnd, |state| {
        (
            state.parent, state.combo_lang, state.combo_open, state.combo_tts_engine, state.combo_voice, state.combo_audio_skip, state.combo_audio_split, state.label_audio_split_text, state.edit_audio_split_text,
            state.checkbox_audio_split_requires_newline, state.checkbox_multilingual, state.button_tts_tuning, state.checkbox_split_on_newline, state.checkbox_word_wrap, state.checkbox_move_cursor, state.checkbox_check_updates
        )
    }) {
        Some(values) => values,
        None => return,
    };

    let settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();
    let labels = options_labels(settings.language);

    let _ = SendMessageW(combo_lang, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let _ = SendMessageW(combo_lang, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(labels.lang_it).as_ptr() as isize));
    let _ = SendMessageW(combo_lang, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(labels.lang_en).as_ptr() as isize));
    let lang_index = match settings.language {
        Language::Italian => 0,
        Language::English => 1,
    };
    let _ = SendMessageW(combo_lang, CB_SETCURSEL, WPARAM(lang_index), LPARAM(0));

    let _ = SendMessageW(combo_open, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let _ = SendMessageW(combo_open, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(labels.open_new_tab).as_ptr() as isize));
    let _ = SendMessageW(combo_open, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(labels.open_new_window).as_ptr() as isize));
    let open_index = match settings.open_behavior {
        OpenBehavior::NewTab => 0,
        OpenBehavior::NewWindow => 1,
    };
    let _ = SendMessageW(combo_open, CB_SETCURSEL, WPARAM(open_index), LPARAM(0));

    let _ = SendMessageW(combo_tts_engine, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let _ = SendMessageW(combo_tts_engine, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(labels.engine_edge).as_ptr() as isize));
    let _ = SendMessageW(combo_tts_engine, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(labels.engine_sapi5).as_ptr() as isize));
    let engine_index = match settings.tts_engine {
        TtsEngine::Edge => 0,
        TtsEngine::Sapi5 => 1,
    };
    let _ = SendMessageW(combo_tts_engine, CB_SETCURSEL, WPARAM(engine_index), LPARAM(0));

    let _ = SendMessageW(checkbox_multilingual, BM_SETCHECK, WPARAM(if settings.tts_only_multilingual { BST_CHECKED.0 as usize } else { 0 }), LPARAM(0));
    let _ = SendMessageW(checkbox_audio_split_requires_newline, BM_SETCHECK, WPARAM(if settings.audiobook_split_text_requires_newline { BST_CHECKED.0 as usize } else { 0 }), LPARAM(0));
    let _ = SendMessageW(checkbox_split_on_newline, BM_SETCHECK, WPARAM(if settings.split_on_newline { BST_CHECKED.0 as usize } else { 0 }), LPARAM(0));
    let _ = SendMessageW(checkbox_word_wrap, BM_SETCHECK, WPARAM(if settings.word_wrap { BST_CHECKED.0 as usize } else { 0 }), LPARAM(0));
    let _ = SendMessageW(checkbox_move_cursor, BM_SETCHECK, WPARAM(if settings.move_cursor_during_reading { BST_CHECKED.0 as usize } else { 0 }), LPARAM(0));
    let _ = SendMessageW(checkbox_check_updates, BM_SETCHECK, WPARAM(if settings.check_updates_on_startup { BST_CHECKED.0 as usize } else { 0 }), LPARAM(0));

    let _ = SendMessageW(combo_audio_skip, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let skip_options = [(10, "10 s"), (30, "30 s"), (60, "1 m"), (120, "2 m"), (300, "5 m")];
    let mut selected_idx = 2;
    for (secs, label) in skip_options.iter() {
        let idx = SendMessageW(combo_audio_skip, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(label).as_ptr() as isize)).0 as usize;
        let _ = SendMessageW(combo_audio_skip, CB_SETITEMDATA, WPARAM(idx), LPARAM(*secs as isize));
        if *secs == settings.audiobook_skip_seconds {
            selected_idx = idx;
        }
    }
    let _ = SendMessageW(combo_audio_skip, CB_SETCURSEL, WPARAM(selected_idx), LPARAM(0));

    let _ = SendMessageW(combo_audio_split, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    let split_options = [
        (0, labels.split_none.to_string()),
        (AUDIOBOOK_SPLIT_BY_TEXT, labels.split_by_text.to_string()),
        (2, format!("2 {}", labels.split_parts)),
        (4, format!("4 {}", labels.split_parts)),
        (6, format!("6 {}", labels.split_parts)),
        (8, format!("8 {}", labels.split_parts)),
    ];
    let mut selected_split_idx = 0;
    for (parts, label) in split_options.iter() {
        let idx = SendMessageW(combo_audio_split, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(label).as_ptr() as isize)).0 as usize;
        let _ = SendMessageW(combo_audio_split, CB_SETITEMDATA, WPARAM(idx), LPARAM(*parts as isize));
        if settings.audiobook_split_by_text && *parts == AUDIOBOOK_SPLIT_BY_TEXT {
            selected_split_idx = idx;
        } else if !settings.audiobook_split_by_text && *parts == settings.audiobook_split {
            selected_split_idx = idx;
        }
    }
    let _ = SendMessageW(combo_audio_split, CB_SETCURSEL, WPARAM(selected_split_idx), LPARAM(0));

    let split_text_wide = to_wide(&settings.audiobook_split_text);
    let _ = SetWindowTextW(edit_audio_split_text, PCWSTR(split_text_wide.as_ptr()));
    update_audio_split_text_visibility(hwnd);

    refresh_voices(hwnd);
}

unsafe fn populate_voice_combo(combo_voice: HWND, voices: &[VoiceInfo], selected: &str, only_multilingual: bool, labels: &OptionsLabels) {
    let _ = SendMessageW(combo_voice, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
    if voices.is_empty() {
        let label = labels.voices_empty; 
        // We could also check if it's loading, but SAPI loads fast. 
        // For Edge, it might be loading.
        // We can check if "loading" logic is needed, but "voices_empty" is safe default.
        let _ = SendMessageW(combo_voice, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(label).as_ptr() as isize));
        let _ = SendMessageW(combo_voice, CB_SETCURSEL, WPARAM(0), LPARAM(0));
        return;
    }
    let mut selected_index: Option<usize> = None;
    let mut combo_index = 0usize;

    for (voice_index, voice) in voices.iter().enumerate() {
        if only_multilingual && !voice.is_multilingual {
            continue;
        }
        let label = format!("{} ({})", voice.short_name, voice.locale);
        let wide = to_wide(&label);
        let idx = SendMessageW(combo_voice, CB_ADDSTRING, WPARAM(0), LPARAM(wide.as_ptr() as isize)).0;
        if idx >= 0 {
            let _ = SendMessageW(combo_voice, CB_SETITEMDATA, WPARAM(idx as usize), LPARAM(voice_index as isize));
            if voice.short_name == selected {
                selected_index = Some(combo_index);
            }
            combo_index += 1;
        }
    }
    
    if let Some(idx) = selected_index {
        let _ = SendMessageW(combo_voice, CB_SETCURSEL, WPARAM(idx), LPARAM(0));
    } else if combo_index > 0 {
         let _ = SendMessageW(combo_voice, CB_SETCURSEL, WPARAM(0), LPARAM(0));
    }
}

unsafe fn apply_options_dialog(hwnd: HWND) {
    let (parent, combo_lang, combo_open, combo_tts_engine, combo_voice, combo_audio_skip, combo_audio_split, edit_audio_split_text, checkbox_audio_split_requires_newline, checkbox_multilingual, _button_tts_tuning, checkbox_split_on_newline, checkbox_word_wrap, checkbox_move_cursor, checkbox_check_updates) = match with_options_state(hwnd, |state| {
        (state.parent, state.combo_lang, state.combo_open, state.combo_tts_engine, state.combo_voice, state.combo_audio_skip, state.combo_audio_split, state.edit_audio_split_text, state.checkbox_audio_split_requires_newline, state.checkbox_multilingual, state.button_tts_tuning, state.checkbox_split_on_newline, state.checkbox_word_wrap, state.checkbox_move_cursor, state.checkbox_check_updates)
    }) { Some(values) => values, None => return };

    let mut settings = with_state(parent, |state| state.settings.clone()).unwrap_or_default();
    let old_language = settings.language;
    let old_word_wrap = settings.word_wrap;
    let (old_engine, old_voice, was_tts_active) = with_state(parent, |state| {
        (state.settings.tts_engine, state.settings.tts_voice.clone(), state.tts_session.is_some())
    }).unwrap_or((settings.tts_engine, settings.tts_voice.clone(), false));

    let lang_sel = SendMessageW(combo_lang, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    settings.language = if lang_sel == 1 { Language::English } else { Language::Italian };

    let open_sel = SendMessageW(combo_open, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    settings.open_behavior = if open_sel == 1 { OpenBehavior::NewWindow } else { OpenBehavior::NewTab };
    
    let engine_sel = SendMessageW(combo_tts_engine, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    settings.tts_engine = if engine_sel == 1 { TtsEngine::Sapi5 } else { TtsEngine::Edge };

    settings.tts_only_multilingual = SendMessageW(checkbox_multilingual, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;
    settings.audiobook_split_text_requires_newline = SendMessageW(checkbox_audio_split_requires_newline, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;
    settings.split_on_newline = SendMessageW(checkbox_split_on_newline, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;
    settings.word_wrap = SendMessageW(checkbox_word_wrap, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;
    settings.move_cursor_during_reading = SendMessageW(checkbox_move_cursor, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;
    settings.check_updates_on_startup = SendMessageW(checkbox_check_updates, BM_GETCHECK, WPARAM(0), LPARAM(0)).0 as u32 == BST_CHECKED.0;

    let voices = with_state(parent, |state| {
        match settings.tts_engine {
            TtsEngine::Edge => state.edge_voices.clone(),
            TtsEngine::Sapi5 => state.sapi_voices.clone(),
        }
    }).unwrap_or_default();
    
    let voice_sel = SendMessageW(combo_voice, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if voice_sel >= 0 {
        let voice_index = SendMessageW(combo_voice, CB_GETITEMDATA, WPARAM(voice_sel as usize), LPARAM(0)).0 as usize;
        if voice_index < voices.len() {
            settings.tts_voice = voices[voice_index].short_name.clone();
        }
    }
    
    let skip_sel = SendMessageW(combo_audio_skip, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if skip_sel >= 0 {
        let skip_secs = SendMessageW(combo_audio_skip, CB_GETITEMDATA, WPARAM(skip_sel as usize), LPARAM(0)).0;
        settings.audiobook_skip_seconds = skip_secs as u32;
    }

    let split_sel = SendMessageW(combo_audio_split, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    if split_sel >= 0 {
        let split_parts = SendMessageW(combo_audio_split, CB_GETITEMDATA, WPARAM(split_sel as usize), LPARAM(0)).0;
        let split_parts = split_parts as u32;
        if split_parts == AUDIOBOOK_SPLIT_BY_TEXT {
            settings.audiobook_split_by_text = true;
            settings.audiobook_split = 0;
        } else {
            settings.audiobook_split_by_text = false;
            settings.audiobook_split = split_parts;
        }
    }

    let text_len = GetWindowTextLengthW(edit_audio_split_text);
    if text_len >= 0 {
        let mut buf = vec![0u16; (text_len + 1) as usize];
        let read = GetWindowTextW(edit_audio_split_text, &mut buf);
        let text = String::from_utf16_lossy(&buf[..read as usize]);
        settings.audiobook_split_text = text;
    }

    let _ = with_state(parent, |state| { state.settings = settings.clone(); });
    let new_language = settings.language;
    save_settings(settings.clone());

    if old_language != new_language {
        rebuild_menus(parent);
    }
    if old_word_wrap != settings.word_wrap {
        apply_word_wrap_to_all_edits(parent, settings.word_wrap);
    }
    refresh_voice_panel(parent);
    if was_tts_active && (old_engine != settings.tts_engine || old_voice != settings.tts_voice) {
        crate::restart_tts_from_current_offset(parent);
    }
    if parent.0 != 0 {
        let _ = PostMessageW(parent, crate::WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
    }
    let _ = DestroyWindow(hwnd);
}

unsafe fn update_audio_split_text_visibility(hwnd: HWND) {
    let (combo_audio_split, label_audio_split_text, edit_audio_split_text, checkbox_audio_split_requires_newline) = match with_options_state(hwnd, |state| {
        (state.combo_audio_split, state.label_audio_split_text, state.edit_audio_split_text, state.checkbox_audio_split_requires_newline)
    }) {
        Some(values) => values,
        None => return,
    };

    let split_sel = SendMessageW(combo_audio_split, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
    let selected = if split_sel >= 0 {
        let split_parts = SendMessageW(combo_audio_split, CB_GETITEMDATA, WPARAM(split_sel as usize), LPARAM(0)).0 as u32;
        split_parts == AUDIOBOOK_SPLIT_BY_TEXT
    } else {
        false
    };

    let show = if selected { SW_SHOW } else { SW_HIDE };
    ShowWindow(label_audio_split_text, show);
    ShowWindow(edit_audio_split_text, show);
    ShowWindow(checkbox_audio_split_requires_newline, show);
    EnableWindow(edit_audio_split_text, selected);
    EnableWindow(checkbox_audio_split_requires_newline, selected);
}

pub(crate) fn ensure_voice_lists_loaded(hwnd: HWND, language: Language) {
    let (has_edge, has_sapi) = unsafe { with_state(hwnd, |state| (!state.edge_voices.is_empty(), !state.sapi_voices.is_empty())) }.unwrap_or((false, false));
    
    if !has_edge {
        thread::spawn(move || {
            match fetch_voice_list() {
                Ok(list) => {
                    let payload = Box::new(list);
                    let _ = unsafe {
                        windows::Win32::UI::WindowsAndMessaging::PostMessageW(
                            hwnd,
                            WM_TTS_VOICES_LOADED,
                            WPARAM(0),
                            LPARAM(Box::into_raw(payload) as isize),
                        )
                    };
                }
                Err(err) => {
                    // Log error but don't show message box for background load unless critical
                    // For now keeping it to avoid spamming user if offline
                    crate::log_debug(&format!("Failed to load Edge voices: {}", err));
                },
            }
        });
    }

    if !has_sapi {
        ensure_sapi_voices_loaded(hwnd, language);
    }
}

fn ensure_sapi_voices_loaded(hwnd: HWND, _language: Language) {
    thread::spawn(move || {
        match crate::sapi5_engine::list_sapi_voices() {
            Ok(list) => {
                 let payload = Box::new(list);
                 let _ = unsafe {
                     windows::Win32::UI::WindowsAndMessaging::PostMessageW(
                         hwnd,
                         WM_TTS_SAPI_VOICES_LOADED,
                         WPARAM(0),
                         LPARAM(Box::into_raw(payload) as isize),
                     )
                 };
            }
            Err(err) => {
                crate::log_debug(&format!("Failed to load SAPI voices: {}", err));
                // Optional: show error if user specifically selected SAPI?
            }
        }
    });
}

fn fetch_voice_list() -> Result<Vec<VoiceInfo>, String> {
    let url = format!("{}?trustedclienttoken={}", VOICE_LIST_URL, TRUSTED_CLIENT_TOKEN);
    let resp = reqwest::blocking::get(url).map_err(|err| err.to_string())?;
    let value: serde_json::Value = resp.json().map_err(|err| err.to_string())?;
    let Some(voices) = value.as_array() else {
        return Err("Risposta non valida".to_string());
    };

    let mut results = Vec::new();
    for voice in voices {
        let short_name = voice["ShortName"].as_str().unwrap_or("").to_string();
        if short_name.is_empty() { continue; }
        let locale = voice["Locale"].as_str().unwrap_or("").to_string();
        let is_multilingual = short_name.contains("Multilingual");
        results.push(VoiceInfo { short_name, locale, is_multilingual });
    }
    results.sort_by(|a, b| a.short_name.cmp(&b.short_name));
    Ok(results)
}

fn focus_language_combo_once(hwnd: HWND) {
    unsafe {
        let (combo, label, ok_button, language) = with_options_state(hwnd, |state| {
            (state.combo_lang, state.label_language, state.ok_button, state.parent)
        }).and_then(|(combo, label, ok_button, parent)| {
            let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
            Some((combo, label, ok_button, language))
        }).unwrap_or((HWND(0), HWND(0), HWND(0), Language::Italian));
        if combo.0 != 0 {
            if label.0 != 0 {
                let labels = options_labels(language);
                let _ = SetWindowTextW(label, PCWSTR(to_wide(" ").as_ptr()));
                let _ = SetWindowTextW(label, PCWSTR(to_wide(labels.label_language).as_ptr()));
            }
            SetForegroundWindow(hwnd);
            if ok_button.0 != 0 {
                SetFocus(ok_button);
            }
            SetFocus(combo);
            let _ = PostMessageW(hwnd, WM_NEXTDLGCTL, WPARAM(combo.0 as usize), LPARAM(1));
            let _ = PostMessageW(combo, WM_SETFOCUS, WPARAM(0), LPARAM(0));
            let sel = SendMessageW(combo, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
            if sel >= 0 {
                let _ = SendMessageW(combo, CB_SETCURSEL, WPARAM(sel as usize), LPARAM(0));
                let _ = SendMessageW(
                    hwnd,
                    WM_COMMAND,
                    WPARAM(OPTIONS_ID_LANG | ((CBN_SELCHANGE as usize) << 16)),
                    LPARAM(combo.0),
                );
            }
        }
    }
}

pub(crate) fn focus_language_combo(hwnd: HWND) {
    focus_language_combo_once(hwnd);
    unsafe {
        let _ = PostMessageW(hwnd, OPTIONS_FOCUS_LANG_MSG, WPARAM(0), LPARAM(0));
        let _ = SetTimer(hwnd, OPTIONS_FOCUS_LANG_TIMER_ID, 80, None);
    }
}
