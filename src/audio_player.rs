use crate::with_state;
use rodio::{Decoder, OutputStream, Sink, Source};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::Duration;
use windows::Win32::Foundation::HWND;

pub struct AudiobookPlayer {
    pub path: PathBuf,
    pub sink: Arc<Sink>,
    pub _stream: OutputStream, // Deve essere mantenuto in vita
    pub is_paused: bool,
    pub start_instant: std::time::Instant,
    pub accumulated_seconds: u64,
    pub volume: f32,
}

pub unsafe fn start_audiobook_playback(hwnd: HWND, path: &Path) {
    let path_buf = path.to_path_buf();

    let bookmark_pos = with_state(hwnd, |state| {
        state
            .bookmarks
            .files
            .get(&path_buf.to_string_lossy().to_string())
            .and_then(|list| list.last()) // Usa l'ultimo segnalibro per l'audio
            .map(|bm| bm.position)
            .unwrap_or(0)
    })
    .unwrap_or(0);

    let hwnd_main = hwnd;
    std::thread::spawn(move || {
        let (_stream, handle) = match OutputStream::try_default() {
            Ok(v) => v,
            Err(_) => return,
        };
        let sink: Arc<Sink> = match Sink::try_new(&handle) {
            Ok(s) => Arc::new(s),
            Err(_) => return,
        };

        let file = match std::fs::File::open(&path_buf) {
            Ok(f) => f,
            Err(_) => return,
        };

        let source: Decoder<_> = match Decoder::new(std::io::BufReader::new(file)) {
            Ok(s) => s,
            Err(_) => return,
        };

        // Salta alla posizione del segnalibro se presente
        if bookmark_pos > 0 {
            let skipped = source.skip_duration(std::time::Duration::from_secs(bookmark_pos as u64));
            sink.append(skipped);
        } else {
            sink.append(source);
        }

        let player = AudiobookPlayer {
            path: path_buf.clone(),
            sink: sink.clone(),
            _stream,
            is_paused: false,
            start_instant: std::time::Instant::now(),
            accumulated_seconds: bookmark_pos as u64,
            volume: 1.0,
        };

        let _ = with_state(hwnd_main, |state| {
            state.active_audiobook = Some(player);
        });
    });
}

pub unsafe fn toggle_audiobook_pause(hwnd: HWND) {
    let _ = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if player.is_paused {
                player.sink.play();
                player.is_paused = false;
                player.start_instant = std::time::Instant::now();
            } else {
                player.sink.pause();
                player.is_paused = true;
                player.accumulated_seconds += player.start_instant.elapsed().as_secs();
            }
        }
    });
}

pub unsafe fn seek_audiobook(hwnd: HWND, seconds: i64) {
    let result = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            if !player.is_paused {
                player.accumulated_seconds += player.start_instant.elapsed().as_secs();
                player.start_instant = std::time::Instant::now();
            }
            let new_pos = (player.accumulated_seconds as i64 + seconds).max(0);
            player.accumulated_seconds = new_pos as u64;
            Some((player.path.clone(), new_pos))
        } else {
            None
        }
    })
    .flatten();

    let (path, current_pos) = match result {
        Some(v) => v,
        None => return,
    };

    stop_audiobook_playback(hwnd);

    let hwnd_main = hwnd;
    std::thread::spawn(move || {
        let (_stream, handle) = OutputStream::try_default().unwrap();
        let sink: Arc<Sink> = Arc::new(Sink::try_new(&handle).unwrap());
        let file = std::fs::File::open(&path).unwrap();
        let source: Decoder<_> = Decoder::new(std::io::BufReader::new(file)).unwrap();

        let skipped = source.skip_duration(Duration::from_secs(current_pos as u64));
        sink.append(skipped);

        let player = AudiobookPlayer {
            path,
            sink: sink.clone(),
            _stream,
            is_paused: false,
            start_instant: std::time::Instant::now(),
            accumulated_seconds: current_pos as u64,
            volume: 1.0,
        };

        let _ = with_state(hwnd_main, |state| {
            state.active_audiobook = Some(player);
        });
    });
}

pub unsafe fn stop_audiobook_playback(hwnd: HWND) {
    let _ = with_state(hwnd, |state| {
        if let Some(player) = state.active_audiobook.take() {
            player.sink.stop();
        }
    });
}

pub unsafe fn start_audiobook_at(hwnd: HWND, path: &Path, seconds: u64) {
    stop_audiobook_playback(hwnd);
    let path_buf = path.to_path_buf();
    let hwnd_main = hwnd;

    std::thread::spawn(move || {
        let (_stream, handle) = match OutputStream::try_default() {
            Ok(v) => v,
            Err(_) => return,
        };
        let sink: Arc<Sink> = match Sink::try_new(&handle) {
            Ok(s) => Arc::new(s),
            Err(_) => return,
        };

        let file = match std::fs::File::open(&path_buf) {
            Ok(f) => f,
            Err(_) => return,
        };

        let source: Decoder<_> = match Decoder::new(std::io::BufReader::new(file)) {
            Ok(s) => s,
            Err(_) => return,
        };

        if seconds > 0 {
            let skipped = source.skip_duration(std::time::Duration::from_secs(seconds));
            sink.append(skipped);
        } else {
            sink.append(source);
        }

        let player = AudiobookPlayer {
            path: path_buf.clone(),
            sink: sink.clone(),
            _stream,
            is_paused: false,
            start_instant: std::time::Instant::now(),
            accumulated_seconds: seconds,
            volume: 1.0,
        };

        let _ = with_state(hwnd_main, |state| {
            state.active_audiobook = Some(player);
        });
    });
}

pub unsafe fn change_audiobook_volume(hwnd: HWND, delta: f32) {
    let _ = with_state(hwnd, |state| {
        if let Some(player) = &mut state.active_audiobook {
            player.volume = (player.volume + delta).clamp(0.0, 1.0);
            player.sink.set_volume(player.volume);
        }
    });
}
