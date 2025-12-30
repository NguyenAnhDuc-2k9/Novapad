use windows::core::PCWSTR;
use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{HBRUSH, COLOR_WINDOW, HFONT};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{WC_BUTTON, WC_COMBOBOXW, WC_STATIC};
use windows::Win32::UI::Input::KeyboardAndMouse::{EnableWindow, GetFocus, SetFocus, VK_ESCAPE, VK_RETURN};
use windows::Win32::UI::WindowsAndMessaging::{
    CallWindowProcW, GetParent, GWLP_WNDPROC, WNDPROC,
    CreateWindowExW, DefWindowProcW, DestroyWindow, GetWindowLongPtrW, LoadCursorW, RegisterClassW,
    SendMessageW, SetForegroundWindow, SetWindowLongPtrW, WM_COMMAND, WM_CREATE, WM_KEYDOWN,
    WM_NCDESTROY, WM_CLOSE, WM_DESTROY, WM_SETFONT, WNDCLASSW, WINDOW_STYLE, WS_CAPTION, WS_CHILD,
    WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT, WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE,
    CB_ADDSTRING, CB_RESETCONTENT, CB_SETCURSEL, CB_GETCURSEL, CB_GETITEMDATA, CB_SETITEMDATA,
    CB_GETDROPPEDSTATE, CREATESTRUCTW, GWLP_USERDATA, HMENU, IDC_ARROW, MSG,
    BS_DEFPUSHBUTTON, CBS_DROPDOWNLIST,
};
use crate::accessibility::{handle_accessibility, to_wide};
use crate::settings::{Language, save_settings};
use crate::{with_state};

const TTS_TUNING_CLASS_NAME: &str = "NovapadTtsTuning";

const TTS_TUNING_ID_SPEED: usize = 9301;
const TTS_TUNING_ID_PITCH: usize = 9302;
const TTS_TUNING_ID_VOLUME: usize = 9303;
const TTS_TUNING_ID_OK: usize = 9304;
const TTS_TUNING_ID_CANCEL: usize = 9305;

struct TtsTuningState {
    parent: HWND,
    owner: HWND,
    combo_speed: HWND,
    combo_pitch: HWND,
    combo_volume: HWND,
    combo_speed_proc: WNDPROC,
    combo_pitch_proc: WNDPROC,
    combo_volume_proc: WNDPROC,
    ok_button: HWND,
}

struct TtsTuningLabels {
    title: &'static str,
    label_speed: &'static str,
    label_pitch: &'static str,
    label_volume: &'static str,
    ok: &'static str,
    cancel: &'static str,
    speed_items: [(&'static str, i32); 11],
    pitch_items: [(&'static str, i32); 11],
    volume_items: [(&'static str, i32); 12],
}

fn tuning_labels(language: Language) -> TtsTuningLabels {
    match language {
        Language::Italian => TtsTuningLabels {
            title: "Tono, velocita' e volume",
            label_speed: "Velocita':",
            label_pitch: "Tono:",
            label_volume: "Volume:",
            ok: "OK",
            cancel: "Annulla",
            speed_items: [
                ("Estremamente lenta", -30),
                ("Molto lenta", -25),
                ("Lenta", -20),
                ("Poco lenta", -10),
                ("Leggermente lenta", -5),
                ("Normale", 0),
                ("Leggermente veloce", 5),
                ("Poco veloce", 10),
                ("Veloce", 15),
                ("Molto veloce", 20),
                ("Velocissima", 30),
            ],
            pitch_items: [
                ("Molto basso", -12),
                ("Basso", -10),
                ("Poco basso", -7),
                ("Leggermente basso", -5),
                ("Poco piu' basso", -2),
                ("Normale", 0),
                ("Poco piu' alto", 2),
                ("Leggermente alto", 5),
                ("Poco alto", 7),
                ("Alto", 9),
                ("Molto alto", 12),
            ],
            volume_items: [
                ("Molto basso", 25),
                ("Basso", 40),
                ("Poco basso", 55),
                ("Medio basso", 70),
                ("Leggermente basso", 85),
                ("Normale", 100),
                ("Leggermente alto", 115),
                ("Medio alto", 130),
                ("Poco alto", 145),
                ("Alto", 160),
                ("Molto alto", 180),
                ("Massimo", 200),
            ],
        },
        Language::English => TtsTuningLabels {
            title: "Pitch, speed, and volume",
            label_speed: "Speed:",
            label_pitch: "Pitch:",
            label_volume: "Volume:",
            ok: "OK",
            cancel: "Cancel",
            speed_items: [
                ("Extremely slow", -30),
                ("Very slow", -25),
                ("Slow", -20),
                ("A bit slow", -10),
                ("Slightly slow", -5),
                ("Normal", 0),
                ("Slightly fast", 5),
                ("A bit fast", 10),
                ("Fast", 15),
                ("Very fast", 20),
                ("Super fast", 30),
            ],
            pitch_items: [
                ("Very low", -12),
                ("Low", -10),
                ("A bit low", -7),
                ("Slightly low", -5),
                ("A little lower", -2),
                ("Normal", 0),
                ("A little higher", 2),
                ("Slightly high", 5),
                ("A bit high", 7),
                ("High", 9),
                ("Very high", 12),
            ],
            volume_items: [
                ("Very low", 25),
                ("Low", 40),
                ("A bit low", 55),
                ("Medium low", 70),
                ("Slightly low", 85),
                ("Normal", 100),
                ("Slightly high", 115),
                ("Medium high", 130),
                ("A bit high", 145),
                ("High", 160),
                ("Very high", 180),
                ("Maximum", 200),
            ],
        },
    }
}

pub unsafe fn handle_navigation(hwnd: HWND, msg: &MSG) -> bool {
    if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
        let focus = GetFocus();
        if GetParent(focus) == hwnd {
            let dropped = SendMessageW(focus, CB_GETDROPPEDSTATE, WPARAM(0), LPARAM(0)).0 != 0;
            if !dropped {
                let _ = with_tts_state(hwnd, |state| {
                    let _ = SendMessageW(hwnd, WM_COMMAND, WPARAM(TTS_TUNING_ID_OK | (0 << 16)), LPARAM(state.ok_button.0));
                });
                return true;
            }
        }
    }
    handle_accessibility(hwnd, msg)
}

pub unsafe fn open(parent: HWND, owner: HWND) {
    let existing = with_state(parent, |state| state.tts_tuning_dialog).unwrap_or(HWND(0));
    if existing.0 != 0 {
        SetForegroundWindow(existing);
        return;
    }
    let hinstance = HINSTANCE(GetModuleHandleW(None).unwrap_or_default().0);
    let class_name = to_wide(TTS_TUNING_CLASS_NAME);
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(LoadCursorW(None, IDC_ARROW).unwrap_or_default().0),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(tts_tuning_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    RegisterClassW(&wc);

    let language = with_state(parent, |state| state.settings.language).unwrap_or_default();
    let labels = tuning_labels(language);

    let params = Box::new(TtsTuningState {
        parent,
        owner,
        combo_speed: HWND(0),
        combo_pitch: HWND(0),
        combo_volume: HWND(0),
        combo_speed_proc: None,
        combo_pitch_proc: None,
        combo_volume_proc: None,
        ok_button: HWND(0),
    });

    let dialog = CreateWindowExW(
        WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
        PCWSTR(class_name.as_ptr()),
        PCWSTR(to_wide(labels.title).as_ptr()),
        WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        0,
        0,
        420,
        260,
        owner,
        None,
        hinstance,
        Some(Box::into_raw(params) as *const std::ffi::c_void),
    );

    if dialog.0 != 0 {
        let _ = with_state(parent, |state| {
            state.tts_tuning_dialog = dialog;
        });
        EnableWindow(owner, false);
        SetForegroundWindow(dialog);
    }
}

unsafe extern "system" fn tts_combo_subclass_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if msg == WM_KEYDOWN && wparam.0 as u32 == VK_RETURN.0 as u32 {
        let parent = GetParent(hwnd);
        if parent.0 != 0 {
            apply_tts_tuning(parent);
            return LRESULT(0);
        }
    }

    let parent = GetParent(hwnd);
    let prev_proc = if parent.0 != 0 {
        with_tts_state(parent, |s| {
            if hwnd == s.combo_speed {
                s.combo_speed_proc
            } else if hwnd == s.combo_pitch {
                s.combo_pitch_proc
            } else if hwnd == s.combo_volume {
                s.combo_volume_proc
            } else {
                None
            }
        }).unwrap_or(None)
    } else {
        None
    };
    if let Some(proc) = prev_proc {
        CallWindowProcW(Some(proc), hwnd, msg, wparam, lparam)
    } else {
        DefWindowProcW(hwnd, msg, wparam, lparam)
    }
}

unsafe extern "system" fn tts_tuning_wndproc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct = lparam.0 as *const CREATESTRUCTW;
            let state_ptr = (*create_struct).lpCreateParams as *mut TtsTuningState;
            if state_ptr.is_null() {
                return LRESULT(0);
            }
            let mut state = Box::from_raw(state_ptr);
            let language = with_state(state.parent, |s| s.settings.language).unwrap_or_default();
            let labels = tuning_labels(language);
            let hfont = with_state(state.parent, |s| s.hfont).unwrap_or(HFONT(0));

            let mut y = 20;
            let label_speed = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(labels.label_speed).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20, y, 140, 20,
                hwnd, HMENU(0), HINSTANCE(0), None,
            );
            let combo_speed = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170, y - 2, 200, 140,
                hwnd, HMENU(TTS_TUNING_ID_SPEED as isize), HINSTANCE(0), None,
            );
            y += 40;

            let label_pitch = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(labels.label_pitch).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20, y, 140, 20,
                hwnd, HMENU(0), HINSTANCE(0), None,
            );
            let combo_pitch = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170, y - 2, 200, 140,
                hwnd, HMENU(TTS_TUNING_ID_PITCH as isize), HINSTANCE(0), None,
            );
            y += 40;

            let label_volume = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(labels.label_volume).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                20, y, 140, 20,
                hwnd, HMENU(0), HINSTANCE(0), None,
            );
            let combo_volume = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_COMBOBOXW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(CBS_DROPDOWNLIST as u32),
                170, y - 2, 200, 140,
                hwnd, HMENU(TTS_TUNING_ID_VOLUME as isize), HINSTANCE(0), None,
            );
            y += 44;

            let ok_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(labels.ok).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                200, y, 80, 28,
                hwnd, HMENU(TTS_TUNING_ID_OK as isize), HINSTANCE(0), None,
            );
            let cancel_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(labels.cancel).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                290, y, 80, 28,
                hwnd, HMENU(TTS_TUNING_ID_CANCEL as isize), HINSTANCE(0), None,
            );

            for ctrl in [label_speed, combo_speed, label_pitch, combo_pitch, label_volume, combo_volume, ok_button, cancel_button] {
                if ctrl.0 != 0 && hfont.0 != 0 {
                    let _ = SendMessageW(ctrl, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

            init_combo(combo_speed, &labels.speed_items);
            init_combo(combo_pitch, &labels.pitch_items);
            init_combo(combo_volume, &labels.volume_items);

            let (rate, pitch, volume) = with_state(state.parent, |s| {
                (s.settings.tts_rate, s.settings.tts_pitch, s.settings.tts_volume)
            }).unwrap_or((0, 0, 100));
            select_combo_value(combo_speed, rate);
            select_combo_value(combo_pitch, pitch);
            select_combo_value(combo_volume, volume);

            state.combo_speed = combo_speed;
            state.combo_pitch = combo_pitch;
            state.combo_volume = combo_volume;
            state.ok_button = ok_button;
            let state_ptr = Box::into_raw(state);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, state_ptr as isize);
            let old_speed = SetWindowLongPtrW(combo_speed, GWLP_WNDPROC, tts_combo_subclass_proc as isize);
            let old_pitch = SetWindowLongPtrW(combo_pitch, GWLP_WNDPROC, tts_combo_subclass_proc as isize);
            let old_volume = SetWindowLongPtrW(combo_volume, GWLP_WNDPROC, tts_combo_subclass_proc as isize);
            let _ = with_tts_state(hwnd, |s| {
                s.combo_speed_proc = std::mem::transmute::<isize, WNDPROC>(old_speed);
                s.combo_pitch_proc = std::mem::transmute::<isize, WNDPROC>(old_pitch);
                s.combo_volume_proc = std::mem::transmute::<isize, WNDPROC>(old_volume);
            });
            SetFocus(combo_speed);
            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xffff) as usize;
            match cmd_id {
                TTS_TUNING_ID_OK => {
                    apply_tts_tuning(hwnd);
                    LRESULT(0)
                }
                TTS_TUNING_ID_CANCEL | 2 => {
                    let _ = DestroyWindow(hwnd);
                    LRESULT(0)
                }
                _ => DefWindowProcW(hwnd, msg, wparam, lparam),
            }
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                apply_tts_tuning(hwnd);
                return LRESULT(0);
            }
            if wparam.0 as u32 == VK_RETURN.0 as u32 {
                let focus = GetFocus();
                let ok = with_tts_state(hwnd, |s| s.ok_button).unwrap_or(HWND(0));
                if focus == ok {
                    apply_tts_tuning(hwnd);
                    return LRESULT(0);
                }
                let is_combo = with_tts_state(hwnd, |s| {
                    focus == s.combo_speed || focus == s.combo_pitch || focus == s.combo_volume
                }).unwrap_or(false);
                if is_combo {
                    apply_tts_tuning(hwnd);
                    return LRESULT(0);
                }
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_CLOSE => {
            let _ = DestroyWindow(hwnd);
            LRESULT(0)
        }
        WM_DESTROY => {
            let (owner, parent) = with_tts_state(hwnd, |s| (s.owner, s.parent)).unwrap_or((HWND(0), HWND(0)));
            if owner.0 != 0 {
                EnableWindow(owner, true);
                crate::app_windows::options_window::focus_language_combo(owner);
            }
            if parent.0 != 0 {
                let _ = with_state(parent, |state| {
                    state.tts_tuning_dialog = HWND(0);
                });
            }
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut TtsTuningState;
            if !ptr.is_null() {
                let state = Box::from_raw(ptr);
                if state.combo_speed.0 != 0 {
                    if let Some(proc) = state.combo_speed_proc {
                        SetWindowLongPtrW(state.combo_speed, GWLP_WNDPROC, proc as isize);
                    }
                }
                if state.combo_pitch.0 != 0 {
                    if let Some(proc) = state.combo_pitch_proc {
                        SetWindowLongPtrW(state.combo_pitch, GWLP_WNDPROC, proc as isize);
                    }
                }
                if state.combo_volume.0 != 0 {
                    if let Some(proc) = state.combo_volume_proc {
                        SetWindowLongPtrW(state.combo_volume, GWLP_WNDPROC, proc as isize);
                    }
                }
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_tts_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut TtsTuningState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut TtsTuningState;
    if ptr.is_null() { None } else { Some(f(&mut *ptr)) }
}

fn init_combo(hwnd: HWND, items: &[(&'static str, i32)]) {
    unsafe {
        let _ = SendMessageW(hwnd, CB_RESETCONTENT, WPARAM(0), LPARAM(0));
        for (label, value) in items {
            let idx = SendMessageW(hwnd, CB_ADDSTRING, WPARAM(0), LPARAM(to_wide(label).as_ptr() as isize)).0 as usize;
            let _ = SendMessageW(hwnd, CB_SETITEMDATA, WPARAM(idx), LPARAM(*value as isize));
        }
        let _ = SendMessageW(hwnd, CB_SETCURSEL, WPARAM(2), LPARAM(0));
    }
}

fn select_combo_value(hwnd: HWND, value: i32) {
    unsafe {
        let count = SendMessageW(hwnd, windows::Win32::UI::WindowsAndMessaging::CB_GETCOUNT, WPARAM(0), LPARAM(0)).0;
        for i in 0..count {
            let data = SendMessageW(hwnd, CB_GETITEMDATA, WPARAM(i as usize), LPARAM(0)).0 as i32;
            if data == value {
                let _ = SendMessageW(hwnd, CB_SETCURSEL, WPARAM(i as usize), LPARAM(0));
                break;
            }
        }
    }
}

unsafe fn apply_tts_tuning(hwnd: HWND) {
    let (parent, _owner, combo_speed, combo_pitch, combo_volume) = match with_tts_state(hwnd, |s| {
        (s.parent, s.owner, s.combo_speed, s.combo_pitch, s.combo_volume)
    }) {
        Some(values) => values,
        None => return,
    };

    let rate = combo_value(combo_speed);
    let pitch = combo_value(combo_pitch);
    let volume = combo_value(combo_volume);

    let mut changed = false;
    let was_tts_active = with_state(parent, |state| state.tts_session.is_some()).unwrap_or(false);
    let _ = with_state(parent, |state| {
        if state.settings.tts_rate != rate {
            state.settings.tts_rate = rate;
            changed = true;
        }
        if state.settings.tts_pitch != pitch {
            state.settings.tts_pitch = pitch;
            changed = true;
        }
        if state.settings.tts_volume != volume {
            state.settings.tts_volume = volume;
            changed = true;
        }
    });
    if changed {
        if let Some(settings) = with_state(parent, |state| state.settings.clone()) {
            save_settings(settings);
        }
        if was_tts_active {
            crate::restart_tts_from_current_offset(parent);
        }
    }

    let _ = DestroyWindow(hwnd);
}

fn combo_value(hwnd: HWND) -> i32 {
    unsafe {
        let sel = SendMessageW(hwnd, CB_GETCURSEL, WPARAM(0), LPARAM(0)).0;
        if sel < 0 {
            return 0;
        }
        SendMessageW(hwnd, CB_GETITEMDATA, WPARAM(sel as usize), LPARAM(0)).0 as i32
    }
}
