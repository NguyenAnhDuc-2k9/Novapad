use std::collections::BTreeMap;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, HBRUSH, HFONT};
use windows::Win32::System::Com::CoTaskMemFree;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Controls::{
    NM_RETURN, NMHDR, NMTVKEYDOWN, PBM_SETPOS, PBM_SETRANGE, PROGRESS_CLASSW, TVN_KEYDOWN,
    WC_BUTTON, WC_EDIT, WC_STATIC,
};
use windows::Win32::UI::Controls::{
    TVGN_CARET, TVI_ROOT, TVIF_PARAM, TVIF_TEXT, TVINSERTSTRUCTW, TVINSERTSTRUCTW_0,
    TVITEMEXW_CHILDREN, TVITEMW, TVM_DELETEITEM, TVM_GETITEMW, TVM_GETNEXTITEM, TVM_INSERTITEMW,
    TVM_SELECTITEM, TVS_HASBUTTONS, TVS_HASLINES, TVS_LINESATROOT, TVS_SHOWSELALWAYS,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{
    EnableWindow, GetFocus, SetFocus, VK_ESCAPE, VK_RETURN,
};
use windows::Win32::UI::Shell::{
    BIF_NEWDIALOGSTYLE, BIF_RETURNONLYFSDIRS, BROWSEINFOW, SHBrowseForFolderW, SHGetPathFromIDListW,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BS_DEFPUSHBUTTON, CreateWindowExW, DefWindowProcW, DestroyWindow, DispatchMessageW,
    FindWindowW, GetMessageW, GetWindowLongPtrW, GetWindowTextLengthW, GetWindowTextW, HMENU,
    IDC_ARROW, IsDialogMessageW, IsWindow, LoadCursorW, MSG, PostMessageW, RegisterClassW,
    SendMessageW, SetForegroundWindow, SetWindowLongPtrW, SetWindowTextW, TranslateMessage,
    WINDOW_STYLE, WM_APP, WM_CLOSE, WM_COMMAND, WM_CREATE, WM_DESTROY, WM_KEYDOWN, WM_NCDESTROY,
    WM_NOTIFY, WM_SETFONT, WNDCLASSW, WS_CAPTION, WS_CHILD, WS_EX_CLIENTEDGE, WS_EX_CONTROLPARENT,
    WS_EX_DLGMODALFRAME, WS_SYSMENU, WS_TABSTOP, WS_VISIBLE, WS_VSCROLL,
};
use windows::core::{PCWSTR, PWSTR, w};

use crate::accessibility::{from_wide, normalize_to_crlf, to_wide};
use crate::file_handler::{
    decode_text, is_doc_path, is_docx_path, is_epub_path, is_html_path, is_mp3_path, is_pdf_path,
    is_spreadsheet_path, read_doc_text, read_docx_text, read_epub_text, read_html_text,
    read_spreadsheet_text,
};
use crate::i18n;
use crate::settings::Language;
use crate::{WM_FOCUS_EDITOR, with_state};

const FIND_IN_FILES_CLASS_NAME: &str = "NovapadFindInFiles";
const FIND_IN_FILES_ID_TERM: usize = 9201;
const FIND_IN_FILES_ID_FOLDER: usize = 9202;
const FIND_IN_FILES_ID_BROWSE: usize = 9203;
const FIND_IN_FILES_ID_SEARCH: usize = 9204;
const FIND_IN_FILES_ID_RESULTS: usize = 9205;
const FIND_IN_FILES_ID_GO: usize = 9206;
const FIND_IN_FILES_ID_PROGRESS: usize = 9207;

const WM_FIND_IN_FILES_PROGRESS: u32 = WM_APP + 40;
const WM_FIND_IN_FILES_DONE: u32 = WM_APP + 41;

struct FindInFilesInit {
    parent: HWND,
    language: Language,
}

struct FindInFilesState {
    hwnd: HWND,
    parent: HWND,
    term_edit: HWND,
    folder_edit: HWND,
    browse_button: HWND,
    search_button: HWND,
    go_button: HWND,
    progress_bar: HWND,
    progress_text: HWND,
    results_tree: HWND,
    results: Vec<SearchResult>,
    language: Language,
    searching: bool,
    cancel_flag: Option<Arc<AtomicBool>>,
}

struct FindInFilesLabels {
    title: String,
    term_label: String,
    folder_label: String,
    browse: String,
    search: String,
    go: String,
    results: String,
    empty_term: String,
    empty_folder: String,
    invalid_folder: String,
    progress: String,
}

#[derive(Clone)]
pub(crate) struct SearchResult {
    path: PathBuf,
    start_utf16: i32,
    len_utf16: i32,
    line: usize,
    snippet: String,
}

#[derive(Clone)]
pub(crate) struct FindInFilesCache {
    pub term: String,
    pub folder: String,
    pub results: Vec<SearchResult>,
}

fn labels(language: Language) -> FindInFilesLabels {
    FindInFilesLabels {
        title: i18n::tr(language, "find_in_files.title"),
        term_label: i18n::tr(language, "find_in_files.term_label"),
        folder_label: i18n::tr(language, "find_in_files.folder_label"),
        browse: i18n::tr(language, "find_in_files.browse"),
        search: i18n::tr(language, "find_in_files.search"),
        go: i18n::tr(language, "find_in_files.go"),
        results: i18n::tr(language, "find_in_files.results"),
        empty_term: i18n::tr(language, "find_in_files.empty_term"),
        empty_folder: i18n::tr(language, "find_in_files.empty_folder"),
        invalid_folder: i18n::tr(language, "find_in_files.invalid_folder"),
        progress: i18n::tr(language, "find_in_files.progress"),
    }
}

pub fn open_find_in_files_dialog(parent: HWND) {
    let class_name = to_wide(FIND_IN_FILES_CLASS_NAME);
    let existing = unsafe { FindWindowW(PCWSTR(class_name.as_ptr()), PCWSTR::null()) };
    if existing.0 != 0 {
        unsafe {
            SetForegroundWindow(existing);
        }
        return;
    }

    let language =
        unsafe { with_state(parent, |state| state.settings.language) }.unwrap_or_default();
    let hinstance = HINSTANCE(unsafe { GetModuleHandleW(None).unwrap_or_default().0 });
    let wc = WNDCLASSW {
        hCursor: windows::Win32::UI::WindowsAndMessaging::HCURSOR(unsafe {
            LoadCursorW(None, IDC_ARROW).unwrap_or_default().0
        }),
        hInstance: hinstance,
        lpszClassName: PCWSTR(class_name.as_ptr()),
        lpfnWndProc: Some(find_in_files_wndproc),
        hbrBackground: HBRUSH((COLOR_WINDOW.0 + 1) as isize),
        ..Default::default()
    };
    unsafe { RegisterClassW(&wc) };

    let init = Box::new(FindInFilesInit { parent, language });
    let labels = labels(language);
    let title = to_wide(&labels.title);

    let hwnd = unsafe {
        CreateWindowExW(
            WS_EX_CONTROLPARENT | WS_EX_DLGMODALFRAME,
            PCWSTR(class_name.as_ptr()),
            PCWSTR(title.as_ptr()),
            WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
            120,
            120,
            620,
            420,
            parent,
            HMENU(0),
            hinstance,
            Some(Box::into_raw(init) as *const _),
        )
    };

    if hwnd.0 == 0 {
        return;
    }

    unsafe {
        EnableWindow(parent, false);
        SetForegroundWindow(hwnd);
    }

    let mut msg = MSG::default();
    loop {
        if !unsafe { IsWindow(hwnd).as_bool() } {
            break;
        }
        let res = unsafe { GetMessageW(&mut msg, HWND(0), 0, 0) };
        if res.0 == 0 {
            break;
        }
        unsafe {
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_ESCAPE.0 as u32 {
                let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
                continue;
            }
            if msg.message == WM_KEYDOWN && msg.wParam.0 as u32 == VK_RETURN.0 as u32 {
                let _ = with_find_state(hwnd, |state| {
                    let focus = GetFocus();
                    let cmd = if focus == state.results_tree {
                        FIND_IN_FILES_ID_GO
                    } else if focus == state.term_edit || focus == state.folder_edit {
                        FIND_IN_FILES_ID_SEARCH
                    } else if focus == state.browse_button {
                        FIND_IN_FILES_ID_BROWSE
                    } else {
                        FIND_IN_FILES_ID_SEARCH
                    };
                    let _ = PostMessageW(hwnd, WM_COMMAND, WPARAM(cmd), LPARAM(0));
                });
                continue;
            }
            if IsDialogMessageW(hwnd, &mut msg).as_bool() {
                continue;
            }
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    unsafe {
        EnableWindow(parent, true);
        SetForegroundWindow(parent);
    }
}

unsafe extern "system" fn find_in_files_wndproc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let create_struct =
                lparam.0 as *const windows::Win32::UI::WindowsAndMessaging::CREATESTRUCTW;
            let init_ptr = (*create_struct).lpCreateParams as *mut FindInFilesInit;
            if init_ptr.is_null() {
                return LRESULT(0);
            }
            let init = Box::from_raw(init_ptr);
            let labels = labels(init.language);
            let hfont = with_state(init.parent, |state| state.hfont).unwrap_or(HFONT(0));

            let term_label = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.term_label).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                16,
                14,
                160,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let term_edit = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_EDIT,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                16,
                36,
                560,
                24,
                hwnd,
                HMENU(FIND_IN_FILES_ID_TERM as isize),
                HINSTANCE(0),
                None,
            );

            let folder_label = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.folder_label).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                16,
                70,
                160,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let folder_edit = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                WC_EDIT,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                16,
                92,
                440,
                24,
                hwnd,
                HMENU(FIND_IN_FILES_ID_FOLDER as isize),
                HINSTANCE(0),
                None,
            );
            let browse_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.browse).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                464,
                92,
                112,
                24,
                hwnd,
                HMENU(FIND_IN_FILES_ID_BROWSE as isize),
                HINSTANCE(0),
                None,
            );

            let search_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.search).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
                464,
                124,
                112,
                26,
                hwnd,
                HMENU(FIND_IN_FILES_ID_SEARCH as isize),
                HINSTANCE(0),
                None,
            );

            let progress_label = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&format!("{} 0%", &labels.progress)).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                16,
                126,
                220,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );

            let progress_bar = CreateWindowExW(
                Default::default(),
                PROGRESS_CLASSW,
                PCWSTR::null(),
                WS_CHILD | WS_VISIBLE,
                16,
                148,
                560,
                18,
                hwnd,
                HMENU(FIND_IN_FILES_ID_PROGRESS as isize),
                HINSTANCE(0),
                None,
            );
            let _ = SendMessageW(
                progress_bar,
                PBM_SETRANGE,
                WPARAM(0),
                LPARAM(((0u16 as u32) | ((100u16 as u32) << 16)) as isize),
            );

            let results_label = CreateWindowExW(
                Default::default(),
                WC_STATIC,
                PCWSTR(to_wide(&labels.results).as_ptr()),
                WS_CHILD | WS_VISIBLE,
                16,
                176,
                160,
                20,
                hwnd,
                HMENU(0),
                HINSTANCE(0),
                None,
            );
            let results_tree = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                w!("SysTreeView32"),
                PCWSTR::null(),
                WS_CHILD
                    | WS_VISIBLE
                    | WS_TABSTOP
                    | WS_VSCROLL
                    | WINDOW_STYLE(
                        (TVS_HASBUTTONS | TVS_HASLINES | TVS_LINESATROOT | TVS_SHOWSELALWAYS)
                            as u32,
                    ),
                16,
                198,
                560,
                160,
                hwnd,
                HMENU(FIND_IN_FILES_ID_RESULTS as isize),
                HINSTANCE(0),
                None,
            );

            let go_button = CreateWindowExW(
                Default::default(),
                WC_BUTTON,
                PCWSTR(to_wide(&labels.go).as_ptr()),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                464,
                362,
                112,
                26,
                hwnd,
                HMENU(FIND_IN_FILES_ID_GO as isize),
                HINSTANCE(0),
                None,
            );

            for control in [
                term_label,
                term_edit,
                folder_label,
                folder_edit,
                browse_button,
                search_button,
                progress_label,
                progress_bar,
                results_label,
                results_tree,
                go_button,
            ] {
                if control.0 != 0 && hfont.0 != 0 {
                    let _ = SendMessageW(control, WM_SETFONT, WPARAM(hfont.0 as usize), LPARAM(1));
                }
            }

            let _ = SetFocus(term_edit);

            let state = Box::new(FindInFilesState {
                hwnd,
                parent: init.parent,
                term_edit,
                folder_edit,
                browse_button,
                search_button,
                go_button,
                progress_bar,
                progress_text: progress_label,
                results_tree,
                results: Vec::new(),
                language: init.language,
                searching: false,
                cancel_flag: None,
            });
            SetWindowLongPtrW(
                hwnd,
                windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA,
                Box::into_raw(state) as isize,
            );

            let _ = with_find_state(hwnd, |state| {
                apply_cache(state);
            });

            LRESULT(0)
        }
        WM_COMMAND => {
            let cmd_id = (wparam.0 & 0xffff) as usize;
            if cmd_id == FIND_IN_FILES_ID_BROWSE {
                let _ = with_find_state(hwnd, |state| {
                    if let Some(folder) = browse_for_folder(hwnd, state.language) {
                        let wide = to_wide(folder.to_string_lossy().as_ref());
                        let _ = SetWindowTextW(state.folder_edit, PCWSTR(wide.as_ptr()));
                        let _ = SetFocus(state.term_edit);
                    }
                });
                LRESULT(0)
            } else if cmd_id == FIND_IN_FILES_ID_SEARCH {
                let _ = with_find_state(hwnd, |state| {
                    start_search(state);
                });
                LRESULT(0)
            } else if cmd_id == FIND_IN_FILES_ID_GO {
                let _ = with_find_state(hwnd, |state| {
                    open_selected_result(state);
                });
                let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
                LRESULT(0)
            } else {
                DefWindowProcW(hwnd, msg, wparam, lparam)
            }
        }
        WM_KEYDOWN => {
            if wparam.0 as u32 == VK_ESCAPE.0 as u32 {
                let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
                return LRESULT(0);
            }
            if wparam.0 as u32 == VK_RETURN.0 as u32 {
                let _ = with_find_state(hwnd, |state| {
                    let focus = GetFocus();
                    if focus == state.results_tree {
                        open_selected_result(state);
                    } else {
                        start_search(state);
                    }
                });
                return LRESULT(0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_NOTIFY => {
            let hdr = lparam.0 as *const NMHDR;
            if !hdr.is_null() {
                unsafe {
                    if (*hdr).code == NM_RETURN
                        && (*hdr).idFrom as usize == FIND_IN_FILES_ID_RESULTS
                    {
                        let _ = with_find_state(hwnd, |state| {
                            open_selected_result(state);
                        });
                        let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
                        return LRESULT(0);
                    }
                    if (*hdr).code == TVN_KEYDOWN
                        && (*hdr).idFrom as usize == FIND_IN_FILES_ID_RESULTS
                    {
                        let key = (lparam.0 as *const NMTVKEYDOWN).as_ref();
                        if let Some(key) = key
                            && key.wVKey == VK_RETURN.0 as u16
                        {
                            let _ = with_find_state(hwnd, |state| {
                                open_selected_result(state);
                            });
                            let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
                            return LRESULT(0);
                        }
                    }
                }
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_FIND_IN_FILES_PROGRESS => {
            let percent = wparam.0 as u32;
            let _ = with_find_state(hwnd, |state| {
                set_progress(state, percent);
            });
            LRESULT(0)
        }
        WM_FIND_IN_FILES_DONE => {
            let results_ptr = lparam.0 as *mut Vec<SearchResult>;
            if !results_ptr.is_null() {
                let results = unsafe { Box::from_raw(results_ptr) };
                let _ = with_find_state(hwnd, |state| {
                    state.results = *results;
                    state.searching = false;
                    state.cancel_flag = None;
                    set_search_enabled(state, true);
                    set_progress(state, 100);
                    populate_results_tree(state);
                    store_cache(state);
                });
            }
            LRESULT(0)
        }
        WM_CLOSE => {
            let _ = with_find_state(hwnd, |state| {
                if let Some(flag) = &state.cancel_flag {
                    flag.store(true, Ordering::SeqCst);
                }
                store_cache(state);
            });
            let _ = DestroyWindow(hwnd);
            LRESULT(0)
        }
        WM_DESTROY => {
            let _ = with_find_state(hwnd, |state| {
                EnableWindow(state.parent, true);
                SetForegroundWindow(state.parent);
            });
            LRESULT(0)
        }
        WM_NCDESTROY => {
            let ptr =
                GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
                    as *mut FindInFilesState;
            if !ptr.is_null() {
                let _ = Box::from_raw(ptr);
            }
            LRESULT(0)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn with_find_state<F, R>(hwnd: HWND, f: F) -> Option<R>
where
    F: FnOnce(&mut FindInFilesState) -> R,
{
    let ptr = GetWindowLongPtrW(hwnd, windows::Win32::UI::WindowsAndMessaging::GWLP_USERDATA)
        as *mut FindInFilesState;
    if ptr.is_null() {
        None
    } else {
        Some(f(&mut *ptr))
    }
}

fn read_control_text(hwnd: HWND) -> String {
    unsafe {
        let len = GetWindowTextLengthW(hwnd);
        if len == 0 {
            return String::new();
        }
        let mut buf = vec![0u16; (len + 1) as usize];
        GetWindowTextW(hwnd, &mut buf);
        from_wide(buf.as_ptr())
    }
}

fn load_cache(parent: HWND) -> Option<FindInFilesCache> {
    unsafe { with_state(parent, |state| state.find_in_files_cache.clone()) }.unwrap_or(None)
}

fn save_cache(parent: HWND, cache: FindInFilesCache) {
    let _ = unsafe {
        with_state(parent, |state| {
            state.find_in_files_cache = Some(cache);
        })
    };
}

fn apply_cache(state: &mut FindInFilesState) {
    let Some(cache) = load_cache(state.parent) else {
        return;
    };
    if !cache.term.is_empty() {
        let wide = to_wide(&cache.term);
        unsafe {
            let _ = SetWindowTextW(state.term_edit, PCWSTR(wide.as_ptr()));
        }
    }
    if !cache.folder.is_empty() {
        let wide = to_wide(&cache.folder);
        unsafe {
            let _ = SetWindowTextW(state.folder_edit, PCWSTR(wide.as_ptr()));
        }
    }
    if !cache.results.is_empty() {
        state.results = cache.results;
        set_progress(state, 100);
        populate_results_tree(state);
    }
}

fn store_cache(state: &FindInFilesState) {
    let term = read_control_text(state.term_edit).trim().to_string();
    let folder = read_control_text(state.folder_edit).trim().to_string();
    let cache = FindInFilesCache {
        term,
        folder,
        results: state.results.clone(),
    };
    save_cache(state.parent, cache);
}

fn start_search(state: &mut FindInFilesState) {
    if state.searching {
        return;
    }
    let term = read_control_text(state.term_edit).trim().to_string();
    let folder = read_control_text(state.folder_edit).trim().to_string();
    let labels = labels(state.language);

    if term.is_empty() {
        unsafe {
            crate::show_error(state.parent, state.language, &labels.empty_term);
        }
        return;
    }
    if folder.is_empty() {
        unsafe {
            crate::show_error(state.parent, state.language, &labels.empty_folder);
        }
        return;
    }

    let folder_path = PathBuf::from(folder);
    if !folder_path.is_dir() {
        unsafe {
            crate::show_error(state.parent, state.language, &labels.invalid_folder);
        }
        return;
    }

    state.searching = true;
    state.results.clear();
    set_search_enabled(state, false);
    set_progress(state, 0);
    clear_results_tree(state.results_tree);

    let hwnd = state.hwnd;
    let language = state.language;
    let term_norm = normalize_to_crlf(&term);
    let folder_path = folder_path.clone();
    let cancel = Arc::new(AtomicBool::new(false));
    state.cancel_flag = Some(cancel.clone());

    std::thread::spawn(move || {
        let files = collect_search_files(&folder_path);
        let total = files.len().max(1);
        let mut results = Vec::new();
        let term_len_utf16 = term_norm.encode_utf16().count() as i32;

        for (idx, path) in files.iter().enumerate() {
            if cancel.load(Ordering::Relaxed) {
                return;
            }
            if let Some(text) = read_text_for_search(path, language) {
                let normalized = normalize_to_crlf(&text);
                collect_matches(&normalized, &term_norm, term_len_utf16, path, &mut results);
            }
            let percent = ((idx + 1) * 100 / total) as u32;
            let _ = unsafe {
                PostMessageW(
                    hwnd,
                    WM_FIND_IN_FILES_PROGRESS,
                    WPARAM(percent as usize),
                    LPARAM(0),
                )
            };
        }

        if cancel.load(Ordering::Relaxed) {
            return;
        }
        let boxed = Box::new(results);
        let _ = unsafe {
            PostMessageW(
                hwnd,
                WM_FIND_IN_FILES_DONE,
                WPARAM(0),
                LPARAM(Box::into_raw(boxed) as isize),
            )
        };
    });
}

fn open_selected_result(state: &mut FindInFilesState) {
    let Some(idx) = selected_result_index(state.results_tree) else {
        return;
    };
    let Some(result) = state.results.get(idx).cloned() else {
        return;
    };
    unsafe {
        crate::editor_manager::open_document_at_position(
            state.parent,
            &result.path,
            result.start_utf16,
            result.len_utf16,
        );
        let _ = PostMessageW(state.parent, WM_FOCUS_EDITOR, WPARAM(0), LPARAM(0));
    }
}

fn set_search_enabled(state: &mut FindInFilesState, enabled: bool) {
    unsafe {
        EnableWindow(state.term_edit, enabled);
        EnableWindow(state.folder_edit, enabled);
        EnableWindow(state.browse_button, enabled);
        EnableWindow(state.search_button, enabled);
        EnableWindow(state.results_tree, enabled);
        EnableWindow(state.go_button, enabled);
    }
}

fn set_progress(state: &mut FindInFilesState, percent: u32) {
    let percent = percent.min(100);
    unsafe {
        let _ = SendMessageW(
            state.progress_bar,
            PBM_SETPOS,
            WPARAM(percent as usize),
            LPARAM(0),
        );
        let label = format!("{} {}%", labels(state.language).progress, percent);
        let wide = to_wide(&label);
        let _ = SetWindowTextW(state.progress_text, PCWSTR(wide.as_ptr()));
    }
}

fn clear_results_tree(tree: HWND) {
    unsafe {
        let _ = SendMessageW(tree, TVM_DELETEITEM, WPARAM(0), LPARAM(TVI_ROOT.0 as isize));
    }
}

fn populate_results_tree(state: &mut FindInFilesState) {
    clear_results_tree(state.results_tree);
    if state.results.is_empty() {
        unsafe {
            let _ = SetFocus(state.results_tree);
        }
        return;
    }

    let mut grouped: BTreeMap<PathBuf, Vec<usize>> = BTreeMap::new();
    for (idx, result) in state.results.iter().enumerate() {
        grouped.entry(result.path.clone()).or_default().push(idx);
    }

    let mut first_child: Option<windows::Win32::UI::Controls::HTREEITEM> = None;
    for (path, indices) in grouped {
        let parent_text = format!("{} ({})", path.display(), indices.len());
        let parent_item = insert_tree_item(state.results_tree, TVI_ROOT, &parent_text, -1);

        for idx in indices {
            let result = &state.results[idx];
            let snippet = truncate_snippet(&result.snippet, 40);
            let line_prefix = i18n::tr_f(
                state.language,
                "find_in_files.line_prefix",
                &[("line", &result.line.to_string())],
            );
            let label = format!("{line_prefix} {snippet}");
            let child = insert_tree_item(state.results_tree, parent_item, &label, idx as isize);
            if first_child.is_none() {
                first_child = Some(child);
            }
        }
    }

    if let Some(child) = first_child {
        unsafe {
            let _ = SendMessageW(
                state.results_tree,
                TVM_SELECTITEM,
                WPARAM(TVGN_CARET as usize),
                LPARAM(child.0 as isize),
            );
            let _ = SetFocus(state.results_tree);
        }
    }
}

fn insert_tree_item(
    tree: HWND,
    parent: windows::Win32::UI::Controls::HTREEITEM,
    text: &str,
    param: isize,
) -> windows::Win32::UI::Controls::HTREEITEM {
    let mut wide = to_wide(text);
    let item = TVITEMW {
        mask: TVIF_TEXT | TVIF_PARAM,
        hItem: windows::Win32::UI::Controls::HTREEITEM(0),
        state: Default::default(),
        stateMask: Default::default(),
        pszText: PWSTR(wide.as_mut_ptr()),
        cchTextMax: (wide.len().saturating_sub(1)) as i32,
        iImage: 0,
        iSelectedImage: 0,
        cChildren: TVITEMEXW_CHILDREN(0),
        lParam: LPARAM(param),
    };
    let insert = TVINSERTSTRUCTW {
        hParent: parent,
        hInsertAfter: windows::Win32::UI::Controls::HTREEITEM(0),
        Anonymous: TVINSERTSTRUCTW_0 { item },
    };
    let res = unsafe {
        SendMessageW(
            tree,
            TVM_INSERTITEMW,
            WPARAM(0),
            LPARAM(&insert as *const _ as isize),
        )
    };
    windows::Win32::UI::Controls::HTREEITEM(res.0)
}

fn selected_result_index(tree: HWND) -> Option<usize> {
    let hitem = unsafe {
        SendMessageW(
            tree,
            TVM_GETNEXTITEM,
            WPARAM(TVGN_CARET as usize),
            LPARAM(0),
        )
    };
    if hitem.0 == 0 {
        return None;
    }
    let mut item = TVITEMW {
        mask: TVIF_PARAM,
        hItem: windows::Win32::UI::Controls::HTREEITEM(hitem.0),
        ..Default::default()
    };
    let ok = unsafe {
        SendMessageW(
            tree,
            TVM_GETITEMW,
            WPARAM(0),
            LPARAM(&mut item as *mut _ as isize),
        )
    };
    if ok.0 == 0 {
        return None;
    }
    let idx = item.lParam.0;
    if idx < 0 { None } else { Some(idx as usize) }
}

fn collect_search_files(folder: &Path) -> Vec<PathBuf> {
    let mut files = Vec::new();
    let mut stack = vec![folder.to_path_buf()];
    while let Some(dir) = stack.pop() {
        let entries = match std::fs::read_dir(&dir) {
            Ok(entries) => entries,
            Err(_) => continue,
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
                continue;
            }
            if is_pdf_path(&path) || is_mp3_path(&path) {
                continue;
            }
            files.push(path);
        }
    }

    files
}

fn read_text_for_search(path: &Path, language: Language) -> Option<String> {
    if is_docx_path(path) {
        return read_docx_text(path, language).ok();
    }
    if is_doc_path(path) {
        return read_doc_text(path, language).ok();
    }
    if is_epub_path(path) {
        return read_epub_text(path, language).ok();
    }
    if is_html_path(path) {
        return read_html_text(path, language).ok().map(|(text, _)| text);
    }
    if is_spreadsheet_path(path) {
        return read_spreadsheet_text(path, language).ok();
    }
    let bytes = std::fs::read(path).ok()?;
    let (text, _) = decode_text(&bytes, language).ok()?;
    Some(text)
}

fn collect_matches(
    text: &str,
    term: &str,
    term_len_utf16: i32,
    path: &Path,
    out: &mut Vec<SearchResult>,
) {
    if term.is_empty() {
        return;
    }
    let mut start = 0usize;
    while let Some(found) = text[start..].find(term) {
        let byte_index = start + found;
        let utf16_index = text[..byte_index].encode_utf16().count() as i32;
        let line = text[..byte_index].bytes().filter(|b| *b == b'\n').count() + 1;
        let (line_start, line_end) = line_bounds(text, byte_index);
        let snippet = text[line_start..line_end].trim().to_string();

        out.push(SearchResult {
            path: path.to_path_buf(),
            start_utf16: utf16_index,
            len_utf16: term_len_utf16,
            line,
            snippet,
        });

        start = byte_index + term.len();
    }
}

fn line_bounds(text: &str, byte_index: usize) -> (usize, usize) {
    let start = text[..byte_index]
        .rfind('\n')
        .map(|idx| idx + 1)
        .unwrap_or(0);
    let end = text[byte_index..]
        .find('\n')
        .map(|idx| byte_index + idx)
        .unwrap_or_else(|| text.len());
    (start, end)
}

fn truncate_snippet(text: &str, max_chars: usize) -> String {
    if max_chars == 0 {
        return String::new();
    }
    let mut out = String::new();
    for (idx, ch) in text.chars().enumerate() {
        if idx >= max_chars {
            out.push_str("...");
            break;
        }
        out.push(ch);
    }
    out
}

fn browse_for_folder(owner: HWND, language: Language) -> Option<PathBuf> {
    let labels = labels(language);
    let title = to_wide(&labels.folder_label);
    let mut bi = BROWSEINFOW {
        hwndOwner: owner,
        lpszTitle: PCWSTR(title.as_ptr()),
        ulFlags: BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE,
        ..Default::default()
    };

    let pidl = unsafe { SHBrowseForFolderW(&mut bi) };
    if pidl.is_null() {
        return None;
    }
    let mut buffer = [0u16; 260];
    let ok = unsafe { SHGetPathFromIDListW(pidl, &mut buffer).as_bool() };
    unsafe { CoTaskMemFree(Some(pidl as *const _)) };
    if !ok {
        return None;
    }
    let path = unsafe { from_wide(buffer.as_ptr()) };
    if path.is_empty() {
        None
    } else {
        Some(PathBuf::from(path))
    }
}
