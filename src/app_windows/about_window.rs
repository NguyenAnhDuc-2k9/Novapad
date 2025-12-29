use windows::core::PCWSTR;
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{MessageBoxW, MB_OK, MB_ICONINFORMATION};
use crate::with_state;
use crate::settings::Language;
use crate::accessibility::to_wide;

pub unsafe fn show(hwnd: HWND) {
    let language = with_state(hwnd, |state| state.settings.language).unwrap_or_default();
    let message = to_wide(about_message(language));
    let title = to_wide(about_title(language));
    MessageBoxW(
        hwnd,
        PCWSTR(message.as_ptr()),
        PCWSTR(title.as_ptr()),
        MB_OK | MB_ICONINFORMATION,
    );
}

fn about_title(language: Language) -> &'static str {
    match language {
        Language::Italian => "Informazioni sul programma",
        Language::English => "About the program",
    }
}

fn about_message(language: Language) -> &'static str {
    match language {
        Language::Italian => "Novapad è un editor di testo per Windows con supporto multi-formato. Apre file TXT, PDF, DOCX, EPUB e fogli di calcolo; legge il testo con sintesi vocale, crea audiolibri MP3 e riproduce audiolibri. Versione: 0.5.0. Autore: Ambrogio Riili.",
        Language::English => "Novapad is a Windows text editor with multi-format support. It opens TXT, PDF, DOCX, EPUB, and spreadsheet files; reads text with TTS, creates MP3 audiobooks, and plays audiobooks. Version: 0.5.0. Author: Ambrogio Riili.",
    }
}

