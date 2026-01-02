#![allow(clippy::io_other_error)]
use std::fs::File;
use std::io::{self, Write};
use std::path::{Path, PathBuf};
use std::thread;

use serde::Deserialize;
use windows::Win32::Foundation::{CloseHandle, HWND, LPARAM, WPARAM};
use windows::Win32::System::Threading::{OpenProcess, PROCESS_SYNCHRONIZE, WaitForSingleObject};
use windows::Win32::UI::WindowsAndMessaging::{
    IDYES, MB_ICONERROR, MB_ICONINFORMATION, MB_ICONQUESTION, MB_OK, MB_SETFOREGROUND, MB_YESNO,
    MessageBoxW, PostMessageW, WM_CLOSE,
};
use windows::core::PCWSTR;

use crate::accessibility::to_wide;
use crate::i18n;
use crate::log_debug;
use crate::settings::{Language, load_settings};
use crate::with_state;

const REPO_OWNER: &str = "Ambro86";
const REPO_NAME: &str = "Novapad";
const USER_AGENT: &str = "NovapadUpdater";

#[derive(Deserialize)]
struct ReleaseAsset {
    name: String,
    browser_download_url: String,
}

#[derive(Deserialize)]
struct ReleaseInfo {
    tag_name: String,
    assets: Vec<ReleaseAsset>,
}

pub(crate) fn check_for_update(hwnd: HWND, interactive: bool) {
    let language = unsafe { with_state(hwnd, |state| state.settings.language).unwrap_or_default() };
    thread::spawn(move || {
        let current_version = env!("CARGO_PKG_VERSION");
        let latest = match fetch_latest_release() {
            Ok(info) => info,
            Err(err) => {
                log_debug(&format!("Update check failed: {err}"));
                if interactive {
                    show_update_error(language, UpdateError::Network);
                }
                return;
            }
        };

        let latest_version = normalize_version(&latest.tag_name);
        if !is_newer_version(&latest_version, current_version) {
            if interactive {
                show_update_info(language, UpdateInfo::NoUpdate);
            }
            return;
        }

        let Some(asset) = select_portable_asset(&latest.assets) else {
            log_debug("Update check: no portable asset found.");
            if interactive {
                show_update_error(language, UpdateError::NoPortableAsset);
            }
            return;
        };

        if !prompt_update(hwnd, language, current_version, &latest_version) {
            return;
        }

        if let Err(err) = download_and_update(hwnd, &asset.browser_download_url) {
            log_debug(&format!("Update failed: {err}"));
            if interactive {
                show_update_error(language, UpdateError::Download);
            }
        }
    });
}

fn fetch_latest_release() -> Result<ReleaseInfo, String> {
    let url = format!("https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest");
    let client = reqwest::blocking::Client::builder()
        .user_agent(USER_AGENT)
        .build()
        .map_err(|err| err.to_string())?;
    let resp = client
        .get(url)
        .send()
        .map_err(|err| err.to_string())?
        .error_for_status()
        .map_err(|err| err.to_string())?;
    resp.json().map_err(|err| err.to_string())
}

fn select_portable_asset(assets: &[ReleaseAsset]) -> Option<&ReleaseAsset> {
    assets.iter().find(|asset| {
        let name = asset.name.to_lowercase();
        name.ends_with(".exe") && !name.contains("setup") && !name.contains(".msi")
    })
}

fn normalize_version(tag: &str) -> String {
    tag.trim().trim_start_matches('v').to_string()
}

fn parse_version(value: &str) -> Option<(u32, u32, u32)> {
    let mut parts = value.split('.').take(3);
    let major = parts.next()?.parse().ok()?;
    let minor = parts.next().unwrap_or("0").parse().ok()?;
    let patch = parts.next().unwrap_or("0").parse().ok()?;
    Some((major, minor, patch))
}

fn is_newer_version(latest: &str, current: &str) -> bool {
    let Some(latest) = parse_version(latest) else {
        return false;
    };
    let Some(current) = parse_version(current) else {
        return false;
    };
    latest > current
}

fn prompt_update(hwnd: HWND, language: Language, current: &str, latest: &str) -> bool {
    let text = i18n::tr_f(
        language,
        "updater.prompt",
        &[("current", current), ("latest", latest)],
    );
    let title = i18n::tr(language, "updater.title");
    let text_w = to_wide(&text);
    let title_w = to_wide(&title);
    let result = unsafe {
        MessageBoxW(
            hwnd,
            PCWSTR(text_w.as_ptr()),
            PCWSTR(title_w.as_ptr()),
            MB_YESNO | MB_ICONQUESTION,
        )
    };
    result == IDYES
}

fn download_and_update(hwnd: HWND, url: &str) -> Result<(), String> {
    let current_exe = std::env::current_exe().map_err(|err| err.to_string())?;
    let temp_path = temp_update_path(&current_exe)?;

    download_file(url, &temp_path)?;

    launch_self_updater(&current_exe, &temp_path).map_err(|err| err.to_string())?;

    unsafe {
        let _ = PostMessageW(hwnd, WM_CLOSE, WPARAM(0), LPARAM(0));
    }
    Ok(())
}

fn temp_update_path(current_exe: &Path) -> Result<PathBuf, String> {
    let file_name = current_exe
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| "Invalid executable name".to_string())?;
    let mut path = std::env::temp_dir();
    path.push(format!("{file_name}.update"));
    Ok(path)
}

fn download_file(url: &str, target: &Path) -> Result<(), String> {
    let client = reqwest::blocking::Client::builder()
        .user_agent(USER_AGENT)
        .build()
        .map_err(|err| err.to_string())?;
    let mut resp = client.get(url).send().map_err(|err| err.to_string())?;
    resp = resp.error_for_status().map_err(|err| err.to_string())?;

    let mut file = File::create(target).map_err(|err| err.to_string())?;
    io::copy(&mut resp, &mut file).map_err(|err| err.to_string())?;
    file.flush().map_err(|err| err.to_string())?;
    Ok(())
}

fn launch_self_updater(current_exe: &Path, new_exe: &Path) -> io::Result<()> {
    let pid = std::process::id();
    let updater_exe = temp_updater_path(current_exe, pid)?;
    std::fs::copy(current_exe, &updater_exe)?;
    std::process::Command::new(updater_exe)
        .arg("--self-update")
        .arg("--pid")
        .arg(pid.to_string())
        .arg("--current")
        .arg(current_exe)
        .arg("--new")
        .arg(new_exe)
        .arg("--restart")
        .spawn()?;
    Ok(())
}

fn temp_updater_path(current_exe: &Path, pid: u32) -> io::Result<PathBuf> {
    let file_name = current_exe
        .file_name()
        .and_then(|name| name.to_str())
        .ok_or_else(|| io::Error::new(io::ErrorKind::Other, "Invalid executable name"))?;
    let mut path = std::env::temp_dir();
    path.push(format!("{file_name}.updater.{pid}.exe"));
    Ok(path)
}

pub(crate) fn run_self_update(args: &[String]) -> Result<(), String> {
    let mut pid: Option<u32> = None;
    let mut current: Option<PathBuf> = None;
    let mut new: Option<PathBuf> = None;
    let mut restart = false;

    let mut it = args.iter().peekable();
    while let Some(arg) = it.next() {
        match arg.as_str() {
            "--pid" => {
                let Some(value) = it.next() else {
                    return Err("Missing --pid value".to_string());
                };
                pid = value.parse().ok();
            }
            "--current" => {
                let Some(value) = it.next() else {
                    return Err("Missing --current value".to_string());
                };
                current = Some(PathBuf::from(value));
            }
            "--new" => {
                let Some(value) = it.next() else {
                    return Err("Missing --new value".to_string());
                };
                new = Some(PathBuf::from(value));
            }
            "--restart" => restart = true,
            _ => {}
        }
    }

    let pid = pid.ok_or_else(|| "Missing pid".to_string())?;
    let current = current.ok_or_else(|| "Missing current path".to_string())?;
    let new = new.ok_or_else(|| "Missing new path".to_string())?;

    wait_for_process_exit(pid);
    if let Err(err) = replace_executable(&current, &new) {
        if err.kind() == io::ErrorKind::PermissionDenied {
            let language = load_settings().language;
            show_permission_error(language);
        } else {
            let language = load_settings().language;
            show_update_error(language, UpdateError::Replace);
        }
        return Err(err.to_string());
    }

    let language = load_settings().language;
    show_update_info(language, UpdateInfo::Completed);
    if restart {
        let _ = std::process::Command::new(&current).spawn();
    }
    Ok(())
}

fn wait_for_process_exit(pid: u32) {
    unsafe {
        if let Ok(handle) = OpenProcess(PROCESS_SYNCHRONIZE, false, pid) {
            let _ = WaitForSingleObject(handle, 20_000);
            let _ = CloseHandle(handle);
        }
    }
}

fn replace_executable(current: &Path, new: &Path) -> Result<(), io::Error> {
    let mut last_err: Option<io::Error> = None;
    for _ in 0..60 {
        match std::fs::rename(new, current) {
            Ok(()) => return Ok(()),
            Err(err) => {
                if err.kind() == io::ErrorKind::PermissionDenied {
                    return Err(err);
                }
                if last_err.is_none() {
                    last_err = Some(err);
                }
            }
        }
        if let Err(err) = std::fs::remove_file(current) {
            if err.kind() == io::ErrorKind::PermissionDenied {
                return Err(err);
            }
            if last_err.is_none() {
                last_err = Some(err);
            }
        }
        match std::fs::rename(new, current) {
            Ok(()) => return Ok(()),
            Err(err) => {
                if err.kind() == io::ErrorKind::PermissionDenied {
                    return Err(err);
                }
                if last_err.is_none() {
                    last_err = Some(err);
                }
            }
        }
        std::thread::sleep(std::time::Duration::from_millis(200));
    }
    Err(last_err
        .unwrap_or_else(|| io::Error::new(io::ErrorKind::Other, "Failed to replace executable")))
}

fn show_permission_error(language: Language) {
    let text = i18n::tr(language, "updater.permission_error");
    let title = i18n::tr(language, "updater.title");
    let text_w = to_wide(&text);
    let title_w = to_wide(&title);
    unsafe {
        MessageBoxW(
            HWND(0),
            PCWSTR(text_w.as_ptr()),
            PCWSTR(title_w.as_ptr()),
            MB_OK | MB_ICONERROR | MB_SETFOREGROUND,
        );
    }
}

enum UpdateError {
    Network,
    NoPortableAsset,
    Download,
    Replace,
}

enum UpdateInfo {
    NoUpdate,
    Completed,
}

fn show_update_error(language: Language, error: UpdateError) {
    let text_key = match error {
        UpdateError::Network => "updater.error.network",
        UpdateError::NoPortableAsset => "updater.error.no_portable",
        UpdateError::Download => "updater.error.download",
        UpdateError::Replace => "updater.error.replace",
    };
    let text = i18n::tr(language, text_key);
    let title = i18n::tr(language, "updater.title");
    let text_w = to_wide(&text);
    let title_w = to_wide(&title);
    unsafe {
        MessageBoxW(
            HWND(0),
            PCWSTR(text_w.as_ptr()),
            PCWSTR(title_w.as_ptr()),
            MB_OK | MB_ICONERROR | MB_SETFOREGROUND,
        );
    }
}

fn show_update_info(language: Language, info: UpdateInfo) {
    let text_key = match info {
        UpdateInfo::NoUpdate => "updater.info.no_update",
        UpdateInfo::Completed => "updater.info.completed",
    };
    let text = i18n::tr(language, text_key);
    let title = i18n::tr(language, "updater.title");
    let text_w = to_wide(&text);
    let title_w = to_wide(&title);
    unsafe {
        MessageBoxW(
            HWND(0),
            PCWSTR(text_w.as_ptr()),
            PCWSTR(title_w.as_ptr()),
            MB_OK | MB_ICONINFORMATION | MB_SETFOREGROUND,
        );
    }
}
