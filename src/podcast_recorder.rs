use crate::accessibility::from_wide;
use crate::audio_capture::{self, AudioRecorderHandle as AudioRecorder};
use crate::graphics_capture::list_monitors;
use crate::mf_encoder;
use crate::settings;
use crate::settings::{PODCAST_DEVICE_DEFAULT, PodcastFormat};
use crate::video_recorder::{self, VideoRecorderHandle};
use chrono::Local;
use std::collections::VecDeque;
use std::fs::{File, OpenOptions};
use std::io::{Seek, SeekFrom, Write};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, AtomicU32, Ordering};
use std::sync::{Arc, Condvar, Mutex};
use std::thread::{self, JoinHandle};
use std::time::{Duration, Instant};
use windows::Win32::Devices::FunctionDiscovery::PKEY_Device_FriendlyName;
use windows::Win32::Foundation::RPC_E_CHANGED_MODE;
use windows::Win32::Media::Audio::{
    AUDCLNT_BUFFERFLAGS_SILENT, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_LOOPBACK,
    DEVICE_STATE_ACTIVE, EDataFlow, IAudioCaptureClient, IAudioClient, IMMDevice,
    IMMDeviceCollection, IMMDeviceEnumerator, MMDeviceEnumerator, WAVEFORMATEX,
    WAVEFORMATEXTENSIBLE, eCapture, eConsole, eRender,
};
use windows::Win32::Media::KernelStreaming::WAVE_FORMAT_EXTENSIBLE;
use windows::Win32::Media::Multimedia::{KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, WAVE_FORMAT_IEEE_FLOAT};
use windows::Win32::System::Com::StructuredStorage::PropVariantToStringAlloc;
use windows::Win32::System::Com::{
    CLSCTX_ALL, COINIT_MULTITHREADED, CoCreateInstance, CoInitializeEx, CoTaskMemFree,
    CoUninitialize, STGM_READ,
};
use windows::Win32::System::Power::{ES_CONTINUOUS, ES_SYSTEM_REQUIRED, SetThreadExecutionState};
use windows::Win32::UI::Shell::PropertiesSystem::IPropertyStore;
use windows::core::PCWSTR;

const TARGET_SAMPLE_RATE: u32 = 44100;
const TARGET_CHANNELS: u16 = 2;
const TARGET_BITS: u16 = 16;
const MIX_CHUNK_FRAMES: usize = 512;

#[derive(Clone)]
pub struct AudioDevice {
    pub id: String,
    pub name: String,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum SampleFormat {
    I16,
    F32,
}

struct DeviceEnumerator {
    _init: ComInit,
    inner: IMMDeviceEnumerator,
}

impl DeviceEnumerator {
    fn new() -> Result<Self, String> {
        let init = ComInit::new()?;
        let inner: IMMDeviceEnumerator = unsafe {
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)
                .map_err(|e| format!("MMDeviceEnumerator failed: {e}"))?
        };
        Ok(Self { _init: init, inner })
    }
}

struct ComInit {
    should_uninit: bool,
}

impl ComInit {
    fn new() -> Result<Self, String> {
        unsafe {
            let result = CoInitializeEx(None, COINIT_MULTITHREADED);
            if let Err(err) = result.ok() {
                if err.code() == RPC_E_CHANGED_MODE {
                    return Ok(Self {
                        should_uninit: false,
                    });
                }
                return Err(format!("CoInitializeEx failed: {err}"));
            }
        }
        Ok(Self {
            should_uninit: true,
        })
    }
}

impl Drop for ComInit {
    fn drop(&mut self) {
        if self.should_uninit {
            unsafe {
                CoUninitialize();
            }
        }
    }
}

pub fn list_input_devices() -> Result<Vec<AudioDevice>, String> {
    list_devices(eCapture)
}

pub fn list_output_devices() -> Result<Vec<AudioDevice>, String> {
    list_devices(eRender)
}

pub fn probe_device(device_id: &str, loopback: bool) -> Result<(), String> {
    let _com = ComInit::new()?;
    let device = resolve_device(device_id, loopback)?;
    let client: IAudioClient = unsafe {
        device
            .Activate(CLSCTX_ALL, None)
            .map_err(|e| format!("AudioClient activate failed: {e}"))?
    };
    let mix_format = unsafe {
        client
            .GetMixFormat()
            .map_err(|e| format!("GetMixFormat failed: {e}"))?
    };
    let mut stream_flags = 0;
    if loopback {
        stream_flags |= AUDCLNT_STREAMFLAGS_LOOPBACK;
    }
    unsafe {
        client
            .Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                stream_flags,
                10_000_000,
                0,
                mix_format,
                None,
            )
            .map_err(|e| format!("AudioClient initialize failed: {e}"))?;
        CoTaskMemFree(Some(mix_format as *const _));
    }
    Ok(())
}

fn list_devices(flow: EDataFlow) -> Result<Vec<AudioDevice>, String> {
    let enumerator = DeviceEnumerator::new()?;
    let collection: IMMDeviceCollection = unsafe {
        enumerator
            .inner
            .EnumAudioEndpoints(flow, DEVICE_STATE_ACTIVE)
            .map_err(|e| format!("EnumAudioEndpoints failed: {e}"))?
    };
    let count = unsafe {
        collection
            .GetCount()
            .map_err(|e| format!("GetCount failed: {e}"))?
    };
    let mut devices = Vec::new();
    for index in 0..count {
        let device: IMMDevice = unsafe {
            collection
                .Item(index)
                .map_err(|e| format!("Device Item failed: {e}"))?
        };
        if let Some(info) = device_info(&device) {
            devices.push(info);
        }
    }
    Ok(devices)
}

fn device_id(device: &IMMDevice) -> Option<String> {
    unsafe {
        let id = device.GetId().ok()?;
        if id.is_null() {
            return None;
        }
        let value = from_wide(id.0);
        CoTaskMemFree(Some(id.0 as *const _));
        if value.is_empty() { None } else { Some(value) }
    }
}

fn device_info(device: &IMMDevice) -> Option<AudioDevice> {
    let id = device_id(device)?;
    let name = device_friendly_name(device).unwrap_or_else(|| id.clone());
    Some(AudioDevice { id, name })
}

fn device_friendly_name(device: &IMMDevice) -> Option<String> {
    unsafe {
        let store: IPropertyStore = device.OpenPropertyStore(STGM_READ).ok()?;
        let value = store.GetValue(&PKEY_Device_FriendlyName).ok()?;
        let name_ptr = PropVariantToStringAlloc(&value).ok()?;
        if name_ptr.is_null() {
            return None;
        }
        let name = from_wide(name_ptr.0);
        CoTaskMemFree(Some(name_ptr.0 as *const _));
        if name.is_empty() { None } else { Some(name) }
    }
}

#[derive(Clone)]
pub struct RecorderConfig {
    pub include_mic: bool,
    pub mic_device_id: String,
    pub mic_gain: f32,
    pub include_system: bool,
    pub system_device_id: String,
    pub system_gain: f32,
    pub include_video: bool,
    pub monitor_id: String,
    pub output_format: PodcastFormat,
    pub mp3_bitrate: u32,
    pub save_folder: PathBuf,
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum RecorderStatus {
    Idle,
    Recording,
    Paused,
    Saving,
    Error,
}

pub struct RecorderHandle {
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
    threads: Vec<JoinHandle<Result<(), String>>>,
    output_path: PathBuf,
    temp_wav: PathBuf,
    temp_mp3: PathBuf,
    format: PodcastFormat,
    // For video recording (integrated like screen_recorder)
    video_recorder: Option<VideoRecorderHandle>,
    system_audio_recorder: Option<AudioRecorder>,
    mic_audio_recorder: Option<AudioRecorder>,
    encoder_stop: Option<Arc<AtomicBool>>,
    encoder_thread: Option<JoinHandle<Result<(), String>>>,
    is_video_recording: bool,
}

struct SharedState {
    status: Mutex<RecorderStatus>,
    last_error: Mutex<Option<String>>,
    started_at: Mutex<Option<Instant>>,
    paused_at: Mutex<Option<Instant>>,
    paused_total: Mutex<Duration>,
    mic_peak: AtomicU32,
    system_peak: AtomicU32,
    include_mic: bool,
    include_system: bool,
}

impl SharedState {
    fn new(include_mic: bool, include_system: bool) -> Self {
        SharedState {
            status: Mutex::new(RecorderStatus::Idle),
            last_error: Mutex::new(None),
            started_at: Mutex::new(None),
            paused_at: Mutex::new(None),
            paused_total: Mutex::new(Duration::ZERO),
            mic_peak: AtomicU32::new(0),
            system_peak: AtomicU32::new(0),
            include_mic,
            include_system,
        }
    }
}

pub struct LevelSnapshot {
    pub mic_peak: u32,
    pub system_peak: u32,
}

pub fn start_recording(config: RecorderConfig) -> Result<RecorderHandle, String> {
    if !config.include_mic && !config.include_system && !config.include_video {
        return Err("No sources selected.".to_string());
    }

    let output_folder = if config.save_folder.as_os_str().is_empty() {
        PathBuf::from(settings::default_podcast_save_folder())
    } else {
        config.save_folder.clone()
    };
    if let Some(parent) = output_folder.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    let _ = std::fs::create_dir_all(&output_folder);

    let timestamp = Local::now().format("%Y-%m-%d_%H-%M-%S").to_string();
    let base_name = format!("Podcast_{timestamp}");

    // If video is included, output will be MP4
    let output_path = if config.include_video {
        output_folder.join(format!("{base_name}.mp4"))
    } else {
        output_folder.join(format!(
            "{}.{}",
            base_name,
            match config.output_format {
                PodcastFormat::Mp3 => "mp3",
                PodcastFormat::Wav => "wav",
            }
        ))
    };
    let temp_wav = output_folder.join(format!("{base_name}.wav.tmp"));
    let temp_mp3 = output_folder.join(format!("{base_name}_tmp.mp3"));

    // Handle video recording path - COPIED FROM screen_recorder.rs
    if config.include_video {
        crate::log_debug(&format!(
            "Starting podcast video recording: {:?}",
            output_path
        ));

        // Find the monitor
        let monitors = list_monitors().unwrap_or_default();
        let monitor = monitors
            .iter()
            .find(|m| m.id == config.monitor_id)
            .or_else(|| monitors.first())
            .ok_or_else(|| "No monitors available for video recording".to_string())?;

        // Start video capture (EXACTLY like screen_recorder)
        let video_recorder = video_recorder::start_video_recording(monitor)?;
        crate::log_debug("Video recorder started");

        // Start system audio capture (EXACTLY like screen_recorder)
        let audio_recorder = audio_capture::start_audio_recording()?;
        crate::log_debug(&format!(
            "System audio recorder started: {} Hz, {} channels",
            audio_recorder.sample_rate, audio_recorder.channels
        ));

        // Start microphone capture if requested
        let mic_recorder = if config.include_mic {
            match start_mic_audio_recording(&config.mic_device_id) {
                Ok(recorder) => {
                    crate::log_debug(&format!(
                        "Mic audio recorder started: {} Hz, {} channels",
                        recorder.sample_rate, recorder.channels
                    ));
                    Some(recorder)
                }
                Err(e) => {
                    crate::log_debug(&format!("Warning: Could not start mic recording: {}", e));
                    None
                }
            }
        } else {
            None
        };

        let mic_gain = config.mic_gain;
        let system_gain = config.system_gain;

        // Create MP4 writer with audio sample rate from capture (EXACTLY like screen_recorder)
        let writer = crate::mf_encoder::Mp4StreamWriter::create(
            &output_path,
            monitor.width,
            monitor.height,
            audio_recorder.sample_rate,
        )?;
        crate::log_debug("MP4 writer created");

        // Start encoder thread (with mic support)
        let encoder_stop = Arc::new(AtomicBool::new(false));
        let encoder_stop_clone = Arc::clone(&encoder_stop);
        let video_queue = Arc::clone(&video_recorder.frame_queue);
        let audio_queue = Arc::clone(&audio_recorder.audio_queue);
        let mic_queue = mic_recorder.as_ref().map(|r| Arc::clone(&r.audio_queue));

        let encoder_thread = thread::spawn(move || {
            podcast_video_encoder_loop(
                writer,
                video_queue,
                audio_queue,
                mic_queue,
                encoder_stop_clone,
                mic_gain,
                system_gain,
            )
        });

        crate::log_debug("Encoder thread started");

        // Return immediately with video recording state
        let shared = Arc::new(SharedState::new(false, false));
        *shared.status.lock().unwrap() = RecorderStatus::Recording;
        *shared.started_at.lock().unwrap() = Some(Instant::now());

        return Ok(RecorderHandle {
            shared,
            stop: Arc::new(AtomicBool::new(false)),
            paused: Arc::new(AtomicBool::new(false)),
            threads: Vec::new(),
            output_path,
            temp_wav,
            temp_mp3,
            format: config.output_format,
            video_recorder: Some(video_recorder),
            system_audio_recorder: Some(audio_recorder),
            mic_audio_recorder: mic_recorder,
            encoder_stop: Some(encoder_stop),
            encoder_thread: Some(encoder_thread),
            is_video_recording: true,
        });
    }

    // Audio-only path
    let shared = Arc::new(SharedState::new(config.include_mic, config.include_system));
    *shared.status.lock().unwrap() = RecorderStatus::Recording;
    *shared.started_at.lock().unwrap() = Some(Instant::now());

    let stop = Arc::new(AtomicBool::new(false));
    let paused = Arc::new(AtomicBool::new(false));

    let mix_buffer = Arc::new(MixBuffer::new());
    let mut threads = Vec::new();

    if config.include_mic {
        crate::log_debug("Starting microphone capture thread");
        let buffer = mix_buffer.clone();
        let shared_state = shared.clone();
        let stop_flag = stop.clone();
        let paused_flag = paused.clone();
        let device_id = config.mic_device_id.clone();
        let mic_gain = config.mic_gain;
        threads.push(thread::spawn(move || {
            crate::log_debug("Microphone capture thread started");
            let result = capture_source(
                SourceKind::Microphone,
                &device_id,
                false,
                mic_gain,
                buffer,
                shared_state.clone(),
                stop_flag.clone(),
                paused_flag,
            );
            if let Err(err) = &result {
                crate::log_debug(&format!("Microphone capture error: {}", err));
                if let Ok(mut error) = shared_state.last_error.lock() {
                    *error = Some(err.clone());
                }
                if let Ok(mut status) = shared_state.status.lock() {
                    *status = RecorderStatus::Error;
                }
                stop_flag.store(true, Ordering::SeqCst);
            } else {
                crate::log_debug("Microphone capture thread stopped normally");
            }
            result
        }));
    }

    if config.include_system {
        crate::log_debug("Starting system audio capture thread");
        let buffer = mix_buffer.clone();
        let shared_state = shared.clone();
        let stop_flag = stop.clone();
        let paused_flag = paused.clone();
        let device_id = config.system_device_id.clone();
        let system_gain = config.system_gain;
        threads.push(thread::spawn(move || {
            crate::log_debug("System audio capture thread started");
            let result = capture_source(
                SourceKind::System,
                &device_id,
                true,
                system_gain,
                buffer,
                shared_state.clone(),
                stop_flag.clone(),
                paused_flag,
            );
            if let Err(err) = &result {
                crate::log_debug(&format!("System audio capture error: {}", err));
                if let Ok(mut error) = shared_state.last_error.lock() {
                    *error = Some(err.clone());
                }
                if let Ok(mut status) = shared_state.status.lock() {
                    *status = RecorderStatus::Error;
                }
                stop_flag.store(true, Ordering::SeqCst);
            } else {
                crate::log_debug("System audio capture thread stopped normally");
            }
            result
        }));
    }

    let keep_awake_stop = stop.clone();
    threads.push(thread::spawn(move || keep_awake_loop(keep_awake_stop)));

    let writer_buffer = mix_buffer.clone();
    let writer_shared = shared.clone();
    let writer_stop = stop.clone();
    let writer_paused = paused.clone();
    let writer_path = match config.output_format {
        PodcastFormat::Mp3 => temp_mp3.clone(),
        PodcastFormat::Wav => temp_wav.clone(),
    };
    let writer_format = config.output_format;
    let writer_bitrate = config.mp3_bitrate;
    threads.push(thread::spawn(move || {
        let result = write_mixed_audio(
            writer_path,
            writer_format,
            writer_bitrate,
            writer_buffer,
            writer_shared.clone(),
            writer_stop.clone(),
            writer_paused,
        );
        if let Err(err) = &result {
            if let Ok(mut error) = writer_shared.last_error.lock() {
                *error = Some(err.clone());
            }
            if let Ok(mut status) = writer_shared.status.lock() {
                *status = RecorderStatus::Error;
            }
            writer_stop.store(true, Ordering::SeqCst);
        }
        result
    }));

    Ok(RecorderHandle {
        shared,
        stop,
        paused,
        threads,
        output_path,
        temp_wav,
        temp_mp3,
        format: config.output_format,
        video_recorder: None,
        system_audio_recorder: None,
        mic_audio_recorder: None,
        encoder_stop: None,
        encoder_thread: None,
        is_video_recording: false,
    })
}

/// Start microphone audio recording using WASAPI (non-loopback)
fn start_mic_audio_recording(device_id: &str) -> Result<AudioRecorder, String> {
    use crate::audio_capture::{AudioQueue, AudioSample};

    let audio_queue = Arc::new(AudioQueue::new(3000));
    let stop = Arc::new(AtomicBool::new(false));

    // Get mic format
    let _com = ComInit::new()?;
    let device = resolve_device(device_id, false)?; // false = not loopback

    let client: IAudioClient = unsafe {
        device
            .Activate(CLSCTX_ALL, None)
            .map_err(|e| format!("AudioClient activate failed: {e}"))?
    };

    let format = unsafe {
        client
            .GetMixFormat()
            .map_err(|e| format!("GetMixFormat failed: {e}"))?
    };

    let sample_rate = unsafe { (*format).nSamplesPerSec };
    let channels = unsafe { (*format).nChannels };

    crate::log_debug(&format!(
        "Mic audio capture will use: {} Hz, {} channels",
        sample_rate, channels
    ));

    let audio_queue_clone = Arc::clone(&audio_queue);
    let stop_clone = Arc::clone(&stop);
    let device_id_clone = device_id.to_string();

    let thread =
        thread::spawn(move || mic_capture_loop(audio_queue_clone, stop_clone, &device_id_clone));

    Ok(audio_capture::create_audio_recorder_handle(
        stop,
        thread,
        audio_queue,
        sample_rate,
        channels,
    ))
}

/// Microphone capture loop - similar to audio_capture but for mic
fn mic_capture_loop(
    audio_queue: Arc<audio_capture::AudioQueue>,
    stop: Arc<AtomicBool>,
    device_id: &str,
) -> Result<(), String> {
    use crate::audio_capture::AudioSample;

    let _com = ComInit::new()?;
    let device = resolve_device(device_id, false)?; // false = not loopback

    let client: IAudioClient = unsafe {
        device
            .Activate(CLSCTX_ALL, None)
            .map_err(|e| format!("AudioClient activate failed: {e}"))?
    };

    let format = unsafe {
        client
            .GetMixFormat()
            .map_err(|e| format!("GetMixFormat failed: {e}"))?
    };

    let sample_rate = unsafe { (*format).nSamplesPerSec };
    let channels = unsafe { (*format).nChannels };

    unsafe {
        client
            .Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                0, // No loopback for mic
                10_000_000,
                0,
                format,
                None,
            )
            .map_err(|e| format!("Initialize failed: {e}"))?;
    }

    let capture_client: IAudioCaptureClient = unsafe {
        client
            .GetService()
            .map_err(|e| format!("GetService failed: {e}"))?
    };

    unsafe {
        client.Start().map_err(|e| format!("Start failed: {e}"))?;
    }

    crate::log_debug("Mic capture loop started");

    while !stop.load(Ordering::SeqCst) {
        thread::sleep(Duration::from_millis(10));

        let packet_length = unsafe {
            capture_client
                .GetNextPacketSize()
                .map_err(|e| format!("GetNextPacketSize failed: {e}"))?
        };

        if packet_length == 0 {
            continue;
        }

        let mut buffer: *mut u8 = std::ptr::null_mut();
        let mut num_frames = 0u32;
        let mut flags = 0u32;

        unsafe {
            capture_client
                .GetBuffer(&mut buffer, &mut num_frames, &mut flags, None, None)
                .map_err(|e| format!("GetBuffer failed: {e}"))?;
        }

        if num_frames > 0 {
            let buffer_size = (num_frames * channels as u32 * 2) as usize;
            let audio_data =
                unsafe { std::slice::from_raw_parts(buffer as *const i16, buffer_size / 2) };

            let sample = AudioSample {
                data: audio_data.to_vec(),
                sample_rate,
                channels,
            };

            audio_queue.push(sample);

            unsafe {
                capture_client
                    .ReleaseBuffer(num_frames)
                    .map_err(|e| format!("ReleaseBuffer failed: {e}"))?;
            }
        }
    }

    unsafe {
        client.Stop().ok();
    }

    crate::log_debug("Mic capture loop stopped");
    Ok(())
}

/// Encoder loop for podcast video recording - with mic mixing support
fn podcast_video_encoder_loop(
    mut writer: crate::mf_encoder::Mp4StreamWriter,
    video_queue: Arc<crate::video_recorder::FrameQueue>,
    system_audio_queue: Arc<audio_capture::AudioQueue>,
    mic_audio_queue: Option<Arc<audio_capture::AudioQueue>>,
    stop: Arc<AtomicBool>,
    mic_gain: f32,
    system_gain: f32,
) -> Result<(), String> {
    crate::log_debug("Encoder loop started");

    let mut last_video_ts = 0i64;
    let mut frames_encoded = 0u64;
    let mut audio_samples_encoded = 0u64;

    loop {
        let should_stop = stop.load(Ordering::SeqCst);

        // If stop signal received, check if we can exit
        if should_stop {
            let video_empty = video_queue.is_empty();
            let system_audio_empty = system_audio_queue.is_empty();
            let mic_audio_empty = mic_audio_queue.as_ref().map_or(true, |q| q.is_empty());

            if video_empty && system_audio_empty && mic_audio_empty {
                thread::sleep(Duration::from_millis(100));

                // Final check
                let video_empty = video_queue.is_empty();
                let system_audio_empty = system_audio_queue.is_empty();
                let mic_audio_empty = mic_audio_queue.as_ref().map_or(true, |q| q.is_empty());

                if video_empty && system_audio_empty && mic_audio_empty {
                    crate::log_debug("Exiting encoder loop - all queues drained");
                    break;
                }
            }
        }

        // Process video frames with adaptive timeout
        // Use shorter timeout if not stopping, to process audio faster
        let video_timeout = if should_stop {
            Duration::from_millis(10)
        } else {
            Duration::from_millis(30) // Reduced from 100ms to prevent audio backlog
        };
        if let Some(frame) = video_queue.pop(video_timeout) {
            if frame.timestamp > last_video_ts {
                match writer.write_video_frame(&frame) {
                    Ok(()) => {
                        last_video_ts = frame.timestamp;
                        frames_encoded += 1;

                        if frames_encoded % 30 == 0 {
                            crate::log_debug(&format!(
                                "Encoded {} video frames, {} audio samples",
                                frames_encoded, audio_samples_encoded
                            ));
                        }
                    }
                    Err(e) => {
                        crate::log_debug(&format!("Error writing video frame: {}", e));
                    }
                }
            }
        }

        // Process audio samples with shorter timeout when stopping
        let audio_timeout = if should_stop {
            Duration::from_millis(5)
        } else {
            Duration::from_millis(10)
        };

        // Check if we should write more audio
        // Don't write if audio is too far behind video (prevents MF blocking)
        let audio_timestamp = writer.get_audio_timestamp();
        let max_audio_lag = 150_000_000; // 15 seconds max lag in 100-nanosecond units (increased for slow encoding)
        let audio_behind_video = last_video_ts.saturating_sub(audio_timestamp);
        let can_write_audio = !should_stop
            || (audio_timestamp <= last_video_ts && audio_behind_video < max_audio_lag);

        if should_stop && !system_audio_queue.is_empty() {
            if !can_write_audio {
                let lag_seconds = audio_behind_video as f64 / 10_000_000.0;
                crate::log_debug(&format!(
                    "Audio too far behind video ({:.2}s lag), discarding remaining audio to prevent MF blocking",
                    lag_seconds
                ));
                // Drain queues without writing
                let mut discarded = 0;
                while let Some(_) = system_audio_queue.pop(Duration::from_millis(1)) {
                    discarded += 1;
                }
                if let Some(mic_queue) = mic_audio_queue.as_ref() {
                    while let Some(_) = mic_queue.pop(Duration::from_millis(1)) {
                        discarded += 1;
                    }
                }
                crate::log_debug(&format!(
                    "Discarded {} audio samples (audio was at {:.2}s, video at {:.2}s)",
                    discarded,
                    audio_timestamp as f64 / 10_000_000.0,
                    last_video_ts as f64 / 10_000_000.0
                ));
            }
        }

        if can_write_audio {
            // Get system audio
            let system_sample = system_audio_queue.pop(audio_timeout);

            // Get mic audio if available
            let mic_sample = if let Some(mic_queue) = mic_audio_queue.as_ref() {
                mic_queue.pop(audio_timeout)
            } else {
                None
            };

            // Mix the two audio streams
            if system_sample.is_some() || mic_sample.is_some() {
                let mixed_data = match (system_sample, mic_sample) {
                    (Some(sys), Some(mic)) => {
                        // System is stereo (2 channels), mic is mono (1 channel)
                        // We need to convert mic mono to stereo and mix
                        let sys_channels = 2;
                        let mic_channels = 1;

                        // Number of frames (samples per channel)
                        let sys_frames = sys.data.len() / sys_channels;
                        let mic_frames = mic.data.len() / mic_channels;
                        let frames = sys_frames.min(mic_frames);

                        let mut mixed = Vec::with_capacity(frames * 2);

                        for frame in 0..frames {
                            let sys_left = sys.data[frame * 2] as f32 * system_gain;
                            let sys_right = sys.data[frame * 2 + 1] as f32 * system_gain;
                            let mic_mono = mic.data[frame] as f32 * mic_gain;

                            // Mix: system (stereo) + mic (converted to stereo by duplicating)
                            let left =
                                ((sys_left + mic_mono) / 2.0).clamp(-32768.0, 32767.0) as i16;
                            let right =
                                ((sys_right + mic_mono) / 2.0).clamp(-32768.0, 32767.0) as i16;

                            mixed.push(left);
                            mixed.push(right);
                        }
                        mixed
                    }
                    (Some(sys), None) => {
                        // Only system audio - apply gain
                        sys.data
                            .iter()
                            .map(|&s| ((s as f32) * system_gain).clamp(-32768.0, 32767.0) as i16)
                            .collect()
                    }
                    (None, Some(mic)) => {
                        // Only mic - convert mono to stereo and apply gain
                        let mut stereo = Vec::with_capacity(mic.data.len() * 2);
                        for &sample in &mic.data {
                            let s = ((sample as f32) * mic_gain).clamp(-32768.0, 32767.0) as i16;
                            stereo.push(s);
                            stereo.push(s); // Duplicate for stereo
                        }
                        stereo
                    }
                    (None, None) => continue,
                };

                match writer.write_audio_samples(&mixed_data) {
                    Ok(()) => {
                        audio_samples_encoded += mixed_data.len() as u64;

                        // Log every second of audio written (only during normal recording)
                        if !should_stop {
                            let audio_duration = writer.get_audio_duration_seconds();
                            if (audio_duration as u64) % 1 == 0 && audio_duration > 0.0 {
                                let prev_duration = ((audio_duration - 0.1) as u64) % 1;
                                if prev_duration != 0 {
                                    let sys_queue_size = system_audio_queue.len();
                                    let mic_queue_size =
                                        mic_audio_queue.as_ref().map_or(0, |q| q.len());
                                    crate::log_debug(&format!(
                                        "Audio written: {:.2}s (sys queue: {}, mic queue: {})",
                                        audio_duration, sys_queue_size, mic_queue_size
                                    ));
                                }
                            }
                        }
                    }
                    Err(e) => {
                        crate::log_debug(&format!("Error writing audio: {}", e));
                    }
                }
            }
        } else if !should_stop {
            // Log when audio writing is blocked during recording (shouldn't happen normally)
            crate::log_debug(&format!(
                "WARNING: Audio writing blocked during recording! audio_ts={}, video_ts={}, lag={}s",
                writer.get_audio_timestamp(),
                last_video_ts,
                (last_video_ts - writer.get_audio_timestamp()) as f64 / 10_000_000.0
            ));
        }
    }

    // Calculate actual durations
    let video_duration_seconds = if frames_encoded > 0 {
        (last_video_ts as f64) / 10_000_000.0
    } else {
        0.0
    };

    let audio_duration_seconds = writer.get_audio_duration_seconds();

    crate::log_debug(&format!(
        "=== RECORDING STATISTICS ===\n\
         Video: {} frames, duration: {:.2} seconds (last_ts: {})\n\
         Audio: {} samples, duration: {:.2} seconds\n\
         ============================",
        frames_encoded,
        video_duration_seconds,
        last_video_ts,
        audio_samples_encoded,
        audio_duration_seconds
    ));

    // Check if any data was written
    if frames_encoded == 0 {
        crate::log_debug(&format!(
            "Warning: No video frames were encoded. Recording may have been too short. \
             Audio samples: {}",
            audio_samples_encoded
        ));
        // Don't return error - allow finalization to attempt completion
    }

    // Finalize the MP4 file
    crate::log_debug("Calling writer.finalize()...");
    writer.finalize()?;
    crate::log_debug("MP4 file finalized successfully");

    Ok(())
}

/// Encode video with audio loop - reads from video queue and audio mix buffer (OLD - not used for video anymore)
fn encode_video_with_audio(
    mut writer: crate::mf_encoder::Mp4StreamWriter,
    video_queue: Arc<crate::video_recorder::FrameQueue>,
    audio_buffer: Arc<MixBuffer>,
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
) -> Result<(), String> {
    crate::log_debug("Encoder loop started");
    let mut last_video_ts = 0i64;
    let mut frames_encoded = 0u64;
    let mut audio_samples_written = 0u64;
    let mut audio_chunks_written = 0u64;

    loop {
        let should_stop = stop.load(Ordering::SeqCst);
        let is_paused = paused.load(Ordering::SeqCst);

        // If stop signal received, check if we can exit
        if should_stop {
            let video_empty = video_queue.is_empty();
            let audio_empty = audio_buffer.is_empty();

            if video_empty && audio_empty {
                thread::sleep(Duration::from_millis(100));

                // Final check
                if video_queue.is_empty() && audio_buffer.is_empty() {
                    crate::log_debug("Exiting encoder loop - all queues drained");
                    break;
                }
            }
        }

        // Process video frames
        if !is_paused {
            let video_timeout = if should_stop {
                Duration::from_millis(10)
            } else {
                Duration::from_millis(30)
            };

            if let Some(frame) = video_queue.pop(video_timeout) {
                if frame.timestamp > last_video_ts {
                    match writer.write_video_frame(&frame) {
                        Ok(()) => {
                            last_video_ts = frame.timestamp;
                            frames_encoded += 1;
                        }
                        Err(e) => {
                            crate::log_debug(&format!("Error writing video frame: {}", e));
                            return Err(format!("Error writing video frame: {}", e));
                        }
                    }
                }
            }
        } else {
            // When paused, still drain video queue to prevent overflow but don't write
            while video_queue.pop(Duration::from_millis(1)).is_some() {}
            thread::sleep(Duration::from_millis(50));
        }

        // Process audio samples
        if !is_paused {
            let audio_timeout = if should_stop {
                Duration::from_millis(5)
            } else {
                Duration::from_millis(10)
            };

            // Check if we should write more audio
            // Don't write if audio is too far behind video (prevents MF blocking)
            let audio_timestamp = writer.get_audio_timestamp();
            let max_audio_lag = 150_000_000; // 15 seconds max lag in 100-nanosecond units
            let audio_behind_video = last_video_ts.saturating_sub(audio_timestamp);
            let can_write_audio = !should_stop
                || (audio_timestamp <= last_video_ts && audio_behind_video < max_audio_lag);

            if should_stop && !audio_buffer.is_empty() {
                if !can_write_audio {
                    let lag_seconds = audio_behind_video as f64 / 10_000_000.0;
                    crate::log_debug(&format!(
                        "Audio too far behind video ({:.2}s lag), discarding remaining audio to prevent MF blocking",
                        lag_seconds
                    ));
                    // Drain queue without writing
                    let mut discarded = 0;
                    while let Some(_) = audio_buffer.pop_chunk(Duration::from_millis(1)) {
                        discarded += 1;
                    }
                    crate::log_debug(&format!(
                        "Discarded {} audio chunks (audio was at {:.2}s, video at {:.2}s)",
                        discarded,
                        audio_timestamp as f64 / 10_000_000.0,
                        last_video_ts as f64 / 10_000_000.0
                    ));
                }
            }

            if can_write_audio {
                match audio_buffer.pop_chunk(audio_timeout) {
                    Some(chunk) => {
                        // Update audio peaks for monitoring
                        let mut max_sample = 0i16;
                        for &sample in &chunk {
                            max_sample = max_sample.max(sample.abs());
                        }
                        let peak = (max_sample as f32 / i16::MAX as f32 * 100.0) as u32;

                        if shared.include_mic && !shared.include_system {
                            shared.mic_peak.store(peak, Ordering::Relaxed);
                        } else if shared.include_system && !shared.include_mic {
                            shared.system_peak.store(peak, Ordering::Relaxed);
                        } else {
                            // Both enabled, split peak evenly
                            shared.mic_peak.store(peak / 2, Ordering::Relaxed);
                            shared.system_peak.store(peak / 2, Ordering::Relaxed);
                        }

                        match writer.write_audio_samples(&chunk) {
                            Ok(()) => {
                                audio_samples_written += chunk.len() as u64;
                                audio_chunks_written += 1;

                                // Log every 100 chunks
                                if audio_chunks_written % 100 == 0 {
                                    crate::log_debug(&format!(
                                        "Audio written: {} chunks, {} samples, {:.2}s",
                                        audio_chunks_written,
                                        audio_samples_written,
                                        writer.get_audio_duration_seconds()
                                    ));
                                }
                            }
                            Err(e) => {
                                crate::log_debug(&format!("Error writing audio: {}", e));
                                return Err(format!("Error writing audio: {}", e));
                            }
                        }
                    }
                    None => {
                        if should_stop && !audio_buffer.is_empty() {
                            crate::log_debug(
                                "WARNING: audio_buffer.pop_chunk() returned None but buffer not empty!",
                            );
                        }
                    }
                }
            } else if !should_stop {
                // Log when audio writing is blocked during recording (shouldn't happen normally)
                crate::log_debug(&format!(
                    "WARNING: Audio writing blocked during recording! audio_ts={}, video_ts={}, lag={}s",
                    writer.get_audio_timestamp(),
                    last_video_ts,
                    (last_video_ts - writer.get_audio_timestamp()) as f64 / 10_000_000.0
                ));
            }
        } else {
            // When paused, drain audio buffer but don't write
            while audio_buffer.pop_chunk(Duration::from_millis(1)).is_some() {}
            thread::sleep(Duration::from_millis(50));
        }
    }

    // Calculate actual durations
    let video_duration_seconds = if frames_encoded > 0 {
        (last_video_ts as f64) / 10_000_000.0
    } else {
        0.0
    };

    let audio_duration_seconds = writer.get_audio_duration_seconds();

    crate::log_debug(&format!(
        "=== RECORDING STATISTICS ===\n\
         Video: {} frames, duration: {:.2} seconds (last_ts: {})\n\
         Audio: {} samples, duration: {:.2} seconds\n\
         ============================",
        frames_encoded,
        video_duration_seconds,
        last_video_ts,
        audio_samples_written,
        audio_duration_seconds
    ));

    // Check if any data was written
    if frames_encoded == 0 {
        crate::log_debug(&format!(
            "Warning: No video frames were encoded. Recording may have been too short. \
             Audio samples: {}",
            audio_samples_written
        ));
        // Don't return error - allow finalization to attempt completion
    }

    // Finalize the MP4 file
    crate::log_debug("Calling writer.finalize()...");
    writer.finalize()?;
    crate::log_debug("MP4 file finalized successfully");

    Ok(())
}

impl RecorderHandle {
    pub fn pause(&self) {
        if !self.paused.swap(true, Ordering::SeqCst) {
            if let Ok(mut paused_at) = self.shared.paused_at.lock() {
                *paused_at = Some(Instant::now());
            }
            if let Ok(mut status) = self.shared.status.lock() {
                *status = RecorderStatus::Paused;
            }
        }
    }

    pub fn resume(&self) {
        if self.paused.swap(false, Ordering::SeqCst) {
            let now = Instant::now();
            if let Ok(mut paused_at) = self.shared.paused_at.lock() {
                if let Some(start) = paused_at.take() {
                    if let Ok(mut total) = self.shared.paused_total.lock() {
                        *total += now.saturating_duration_since(start);
                    }
                }
            }
            if let Ok(mut status) = self.shared.status.lock() {
                *status = RecorderStatus::Recording;
            }
        }
    }

    pub fn stop(self) -> Result<PathBuf, String> {
        self.stop_with_progress(|_| {}, None)
    }

    pub fn stop_with_progress<F>(
        mut self,
        mut progress: F,
        cancel: Option<Arc<AtomicBool>>,
    ) -> Result<PathBuf, String>
    where
        F: FnMut(u32),
    {
        crate::log_debug("Stopping podcast recording");

        // If video recording, use screen_recorder stop logic
        if self.is_video_recording {
            crate::log_debug("Stopping podcast video recording");

            // FIRST: Signal encoder to stop (EXACTLY like screen_recorder)
            if let Some(encoder_stop) = self.encoder_stop.as_ref() {
                encoder_stop.store(true, Ordering::SeqCst);
                crate::log_debug("Signaled encoder to stop");
            }

            // SECOND: Signal recorders to stop (EXACTLY like screen_recorder)
            if let Some(video_rec) = self.video_recorder.as_ref() {
                video_rec.signal_stop();
            }
            if let Some(audio_rec) = self.system_audio_recorder.as_ref() {
                audio_rec.signal_stop();
            }
            if let Some(mic_rec) = self.mic_audio_recorder.as_ref() {
                mic_rec.signal_stop();
            }
            crate::log_debug("Signaled recorders to stop");

            // THIRD: Wait for encoder to finish (EXACTLY like screen_recorder)
            if let Some(thread) = self.encoder_thread.take() {
                crate::log_debug("Waiting for encoder to finish...");
                thread
                    .join()
                    .map_err(|_| "Encoder thread panicked".to_string())??;
                crate::log_debug("Encoder stopped");
            }

            // FOURTH: Wait for recorders to shut down (EXACTLY like screen_recorder)
            if let Some(video_rec) = self.video_recorder.take() {
                video_rec.join()?;
                crate::log_debug("Video recorder stopped");
            }

            if let Some(audio_rec) = self.system_audio_recorder.take() {
                audio_rec.join()?;
                crate::log_debug("System audio recorder stopped");
            }

            if let Some(mic_rec) = self.mic_audio_recorder.take() {
                mic_rec.join()?;
                crate::log_debug("Mic audio recorder stopped");
            }

            crate::log_debug("Podcast video recording stopped successfully");

            if let Ok(mut status) = self.shared.status.lock() {
                *status = RecorderStatus::Idle;
            }
            return Ok(self.output_path.clone());
        }

        // Audio-only path
        // FIRST: Signal encoder to stop (so it knows to drain queues and exit)
        self.stop.store(true, Ordering::SeqCst);
        crate::log_debug("Signaled encoder to stop");

        if let Ok(mut status) = self.shared.status.lock() {
            *status = RecorderStatus::Saving;
        }

        // SECOND: Wait for all threads to finish
        let threads = std::mem::take(&mut self.threads);
        for handle in threads {
            match handle.join() {
                Ok(Ok(())) => {}
                Ok(Err(err)) => {
                    crate::log_debug(&format!("Thread error: {}", err));
                    self.set_error(&err);
                    return Err(err);
                }
                Err(_) => {
                    let err = "Recording thread panicked.".to_string();
                    crate::log_debug(&err);
                    self.set_error(&err);
                    return Err(err);
                }
            }
        }
        crate::log_debug("All threads stopped");

        if let Some(cancel) = cancel.as_ref() {
            if cancel.load(Ordering::Relaxed) {
                let _ = std::fs::remove_file(&self.temp_wav);
                let _ = std::fs::remove_file(&self.temp_mp3);
                if self.is_video_recording {
                    let _ = std::fs::remove_file(&self.output_path);
                }
                return Err("Saving canceled.".to_string());
            }
        }

        // For video recording, the MP4 file is already complete
        if self.is_video_recording {
            progress(100);
        } else if self.format == PodcastFormat::Mp3 {
            progress(100);
            if let Err(err) = rename_atomic(&self.temp_mp3, &self.output_path) {
                self.set_error(&err);
                return Err(err);
            }
        } else {
            progress(100);
            if let Err(err) = rename_atomic(&self.temp_wav, &self.output_path) {
                self.set_error(&err);
                return Err(err);
            }
        }

        if let Ok(mut status) = self.shared.status.lock() {
            *status = RecorderStatus::Idle;
        }
        Ok(self.output_path.clone())
    }

    pub fn status(&self) -> RecorderStatus {
        self.shared
            .status
            .lock()
            .map(|status| *status)
            .unwrap_or(RecorderStatus::Error)
    }

    pub fn levels(&self) -> LevelSnapshot {
        LevelSnapshot {
            mic_peak: self.shared.mic_peak.load(Ordering::Relaxed),
            system_peak: self.shared.system_peak.load(Ordering::Relaxed),
        }
    }

    pub fn elapsed(&self) -> Duration {
        let start = self.shared.started_at.lock().ok().and_then(|s| *s);
        let start = match start {
            Some(value) => value,
            None => return Duration::ZERO,
        };
        let paused_total = self
            .shared
            .paused_total
            .lock()
            .map(|v| *v)
            .unwrap_or(Duration::ZERO);
        let paused_at = self.shared.paused_at.lock().ok().and_then(|s| *s);
        let now = Instant::now();
        let mut elapsed = now.saturating_duration_since(start);
        if let Some(paused_at) = paused_at {
            elapsed = paused_at.saturating_duration_since(start);
        }
        elapsed.saturating_sub(paused_total)
    }

    pub fn take_error(&self) -> Option<String> {
        self.shared.last_error.lock().ok()?.take()
    }

    fn set_error(&self, message: &str) {
        if let Ok(mut err) = self.shared.last_error.lock() {
            *err = Some(message.to_string());
        }
        if let Ok(mut status) = self.shared.status.lock() {
            *status = RecorderStatus::Error;
        }
    }
}

fn rename_atomic(src: &Path, dest: &Path) -> Result<(), String> {
    if dest.exists() {
        let _ = std::fs::remove_file(dest);
    }
    std::fs::rename(src, dest).map_err(|e| e.to_string())
}

fn keep_awake_loop(stop: Arc<AtomicBool>) -> Result<(), String> {
    unsafe {
        let _ = SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED);
    }
    while !stop.load(Ordering::SeqCst) {
        thread::sleep(Duration::from_secs(30));
        unsafe {
            let _ = SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED);
        }
    }
    unsafe {
        let _ = SetThreadExecutionState(ES_CONTINUOUS);
    }
    Ok(())
}

#[derive(Clone, Copy)]
enum SourceKind {
    Microphone,
    System,
}

struct MixBuffer {
    inner: Mutex<MixQueues>,
    condvar: Condvar,
}

struct MixQueues {
    mic: VecDeque<f32>,
    system: VecDeque<f32>,
}

impl MixBuffer {
    fn new() -> Self {
        MixBuffer {
            inner: Mutex::new(MixQueues {
                mic: VecDeque::new(),
                system: VecDeque::new(),
            }),
            condvar: Condvar::new(),
        }
    }

    fn push(&self, source: SourceKind, samples: Vec<f32>) {
        let mut inner = self.inner.lock().unwrap_or_else(|e| e.into_inner());
        match source {
            SourceKind::Microphone => inner.mic.extend(samples),
            SourceKind::System => inner.system.extend(samples),
        }
        self.condvar.notify_one();
    }

    fn pop_chunk(&self, timeout: Duration) -> Option<Vec<i16>> {
        let mut inner = self.inner.lock().unwrap_or_else(|e| e.into_inner());

        // Wait for data if queues are empty
        if inner.mic.is_empty() && inner.system.is_empty() {
            let result = self.condvar.wait_timeout(inner, timeout).unwrap();
            inner = result.0;
        }

        // Determine how many samples we can mix
        let mic_len = inner.mic.len();
        let system_len = inner.system.len();

        if mic_len == 0 && system_len == 0 {
            return None;
        }

        // Take up to MIX_CHUNK_FRAMES samples
        let count = mic_len.max(system_len).min(MIX_CHUNK_FRAMES);
        let mut output = Vec::with_capacity(count);

        for _ in 0..count {
            let mic_sample = inner.mic.pop_front().unwrap_or(0.0);
            let system_sample = inner.system.pop_front().unwrap_or(0.0);

            // Mix and convert to i16
            let mixed = mic_sample + system_sample;
            let clamped = mixed.clamp(-1.0, 1.0);
            let sample_i16 = (clamped * i16::MAX as f32) as i16;
            output.push(sample_i16);
        }

        Some(output)
    }

    fn is_empty(&self) -> bool {
        let inner = self.inner.lock().unwrap_or_else(|e| e.into_inner());
        inner.mic.is_empty() && inner.system.is_empty()
    }
}

fn write_mixed_audio(
    path: PathBuf,
    format: PodcastFormat,
    mp3_bitrate: u32,
    buffer: Arc<MixBuffer>,
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
) -> Result<(), String> {
    match format {
        PodcastFormat::Mp3 => {
            write_mixed_audio_mp3(path, mp3_bitrate, buffer, shared, stop, paused)
        }
        PodcastFormat::Wav => write_mixed_audio_wav(path, buffer, shared, stop, paused),
    }
}

fn write_mixed_audio_wav(
    path: PathBuf,
    buffer: Arc<MixBuffer>,
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
) -> Result<(), String> {
    let mut writer = WavWriter::create(&path)?;
    let mut last_write = Instant::now();
    loop {
        if stop.load(Ordering::SeqCst) {
            break;
        }
        if paused.load(Ordering::SeqCst) {
            thread::sleep(Duration::from_millis(30));
            continue;
        }

        let mut inner = buffer.inner.lock().unwrap_or_else(|e| e.into_inner());
        let (need_mic, need_sys) = (shared.include_mic, shared.include_system);
        let available_mic = inner.mic.len() / TARGET_CHANNELS as usize;
        let available_sys = inner.system.len() / TARGET_CHANNELS as usize;
        let can_mix = if need_mic && need_sys {
            available_mic >= MIX_CHUNK_FRAMES && available_sys >= MIX_CHUNK_FRAMES
        } else if need_mic {
            available_mic >= MIX_CHUNK_FRAMES
        } else {
            available_sys >= MIX_CHUNK_FRAMES
        };

        if !can_mix {
            let _ = buffer
                .condvar
                .wait_timeout(inner, Duration::from_millis(40));
            continue;
        }

        let frames = MIX_CHUNK_FRAMES;
        let mut mixed = Vec::with_capacity(frames * TARGET_CHANNELS as usize);
        for _ in 0..frames {
            let mut left = 0.0f32;
            let mut right = 0.0f32;
            if need_mic {
                left += inner.mic.pop_front().unwrap_or(0.0);
                right += inner.mic.pop_front().unwrap_or(0.0);
            }
            if need_sys {
                left += inner.system.pop_front().unwrap_or(0.0);
                right += inner.system.pop_front().unwrap_or(0.0);
            }
            if need_mic && need_sys {
                left *= 0.5;
                right *= 0.5;
            }
            mixed.push(left.clamp(-1.0, 1.0));
            mixed.push(right.clamp(-1.0, 1.0));
        }
        drop(inner);

        writer.write_f32(&mixed)?;

        let elapsed = last_write.elapsed();
        if elapsed < Duration::from_millis(10) {
            thread::sleep(Duration::from_millis(5));
        }
        last_write = Instant::now();
    }
    writer.finalize()?;
    Ok(())
}

fn write_mixed_audio_mp3(
    path: PathBuf,
    mp3_bitrate: u32,
    buffer: Arc<MixBuffer>,
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
) -> Result<(), String> {
    let mut writer = mf_encoder::Mp3StreamWriter::create(
        &path,
        mp3_bitrate,
        TARGET_SAMPLE_RATE,
        TARGET_CHANNELS,
    )?;
    let mut last_write = Instant::now();
    loop {
        if stop.load(Ordering::SeqCst) {
            break;
        }
        if paused.load(Ordering::SeqCst) {
            thread::sleep(Duration::from_millis(30));
            continue;
        }

        let mut inner = buffer.inner.lock().unwrap_or_else(|e| e.into_inner());
        let (need_mic, need_sys) = (shared.include_mic, shared.include_system);
        let available_mic = inner.mic.len() / TARGET_CHANNELS as usize;
        let available_sys = inner.system.len() / TARGET_CHANNELS as usize;
        let can_mix = if need_mic && need_sys {
            available_mic >= MIX_CHUNK_FRAMES && available_sys >= MIX_CHUNK_FRAMES
        } else if need_mic {
            available_mic >= MIX_CHUNK_FRAMES
        } else {
            available_sys >= MIX_CHUNK_FRAMES
        };

        if !can_mix {
            let _ = buffer
                .condvar
                .wait_timeout(inner, Duration::from_millis(40));
            continue;
        }

        let frames = MIX_CHUNK_FRAMES;
        let mut mixed = Vec::with_capacity(frames * TARGET_CHANNELS as usize);
        for _ in 0..frames {
            let mut left = 0.0f32;
            let mut right = 0.0f32;
            if need_mic {
                left += inner.mic.pop_front().unwrap_or(0.0);
                right += inner.mic.pop_front().unwrap_or(0.0);
            }
            if need_sys {
                left += inner.system.pop_front().unwrap_or(0.0);
                right += inner.system.pop_front().unwrap_or(0.0);
            }
            if need_mic && need_sys {
                left *= 0.5;
                right *= 0.5;
            }
            mixed.push(left.clamp(-1.0, 1.0));
            mixed.push(right.clamp(-1.0, 1.0));
        }
        drop(inner);

        let mut pcm = Vec::with_capacity(mixed.len());
        for sample in mixed {
            let v = (sample.clamp(-1.0, 1.0) * i16::MAX as f32) as i16;
            pcm.push(v);
        }
        writer.write_i16(&pcm)?;

        let elapsed = last_write.elapsed();
        if elapsed < Duration::from_millis(10) {
            thread::sleep(Duration::from_millis(5));
        }
        last_write = Instant::now();
    }
    writer.finalize()?;
    Ok(())
}

fn capture_source(
    kind: SourceKind,
    device_id: &str,
    loopback: bool,
    gain: f32,
    buffer: Arc<MixBuffer>,
    shared: Arc<SharedState>,
    stop: Arc<AtomicBool>,
    paused: Arc<AtomicBool>,
) -> Result<(), String> {
    let kind_name = match kind {
        SourceKind::Microphone => "Microphone",
        SourceKind::System => "System",
    };
    crate::log_debug(&format!("{} capture_source started", kind_name));

    let _com = ComInit::new()?;
    let device = resolve_device(device_id, loopback)?;
    crate::log_debug(&format!("{} device resolved", kind_name));
    let client: IAudioClient = unsafe {
        device
            .Activate(CLSCTX_ALL, None)
            .map_err(|e| format!("AudioClient activate failed: {e}"))?
    };

    let mix_format = unsafe {
        client
            .GetMixFormat()
            .map_err(|e| format!("GetMixFormat failed: {e}"))?
    };
    let (input_rate, input_channels, input_format) = parse_format(unsafe { &*mix_format });

    let mut stream_flags = 0;
    if loopback {
        stream_flags |= AUDCLNT_STREAMFLAGS_LOOPBACK;
    }
    unsafe {
        client
            .Initialize(
                AUDCLNT_SHAREMODE_SHARED,
                stream_flags,
                10_000_000,
                0,
                mix_format,
                None,
            )
            .map_err(|e| format!("AudioClient initialize failed: {e}"))?;
    }
    unsafe {
        CoTaskMemFree(Some(mix_format as *const _));
    }

    let capture: IAudioCaptureClient = unsafe {
        client
            .GetService()
            .map_err(|e| format!("GetService capture failed: {e}"))?
    };
    unsafe {
        client.Start().map_err(|e| format!("Start failed: {e}"))?;
    }

    let mut resampler =
        LinearResampler::new(input_rate, TARGET_SAMPLE_RATE, input_channels as usize);

    loop {
        if stop.load(Ordering::SeqCst) {
            break;
        }
        if paused.load(Ordering::SeqCst) {
            thread::sleep(Duration::from_millis(15));
            continue;
        }

        let mut packet_len = unsafe {
            capture
                .GetNextPacketSize()
                .map_err(|e| format!("GetNextPacketSize failed: {e}"))?
        };
        while packet_len > 0 {
            let mut data_ptr: *mut u8 = std::ptr::null_mut();
            let mut frames = 0u32;
            let mut flags = 0u32;
            unsafe {
                capture
                    .GetBuffer(&mut data_ptr, &mut frames, &mut flags, None, None)
                    .map_err(|e| format!("GetBuffer failed: {e}"))?;
            }
            let samples = if flags & (AUDCLNT_BUFFERFLAGS_SILENT.0 as u32) != 0 {
                vec![0f32; frames as usize * input_channels as usize]
            } else {
                read_samples(data_ptr, frames, input_channels, input_format)
            };
            unsafe {
                capture
                    .ReleaseBuffer(frames)
                    .map_err(|e| format!("ReleaseBuffer failed: {e}"))?;
            }

            update_peak(&shared, &kind, &samples);
            let resampled = resampler.push(&samples);
            let mut stereo = to_stereo(&resampled, input_channels as usize);

            // Apply gain
            if gain != 1.0 {
                for sample in stereo.iter_mut() {
                    *sample = (*sample * gain).clamp(-1.0, 1.0);
                }
            }

            buffer.push(kind, stereo);
            packet_len = unsafe {
                capture
                    .GetNextPacketSize()
                    .map_err(|e| format!("GetNextPacketSize failed: {e}"))?
            };
        }
        thread::sleep(Duration::from_millis(10));
    }

    unsafe {
        let _ = client.Stop();
    }
    Ok(())
}

fn resolve_device(device_id: &str, loopback: bool) -> Result<IMMDevice, String> {
    let enumerator = DeviceEnumerator::new()?;
    if device_id.is_empty() || device_id == PODCAST_DEVICE_DEFAULT {
        let flow = if loopback { eRender } else { eCapture };
        return unsafe {
            enumerator
                .inner
                .GetDefaultAudioEndpoint(flow, eConsole)
                .map_err(|e| format!("GetDefaultAudioEndpoint failed: {e}"))
        };
    }
    let wide = crate::accessibility::to_wide(device_id);
    unsafe {
        enumerator
            .inner
            .GetDevice(PCWSTR(wide.as_ptr()))
            .map_err(|e| format!("GetDevice failed: {e}"))
    }
}

fn parse_format(fmt: &WAVEFORMATEX) -> (u32, u16, SampleFormat) {
    let channels = fmt.nChannels;
    let rate = fmt.nSamplesPerSec;
    let mut format = match fmt.wFormatTag as u32 {
        WAVE_FORMAT_IEEE_FLOAT => SampleFormat::F32,
        _ => SampleFormat::I16,
    };
    if fmt.wFormatTag as u32 == WAVE_FORMAT_EXTENSIBLE {
        let ext = unsafe { &*(fmt as *const _ as *const WAVEFORMATEXTENSIBLE) };
        let subformat = unsafe { std::ptr::read_unaligned(std::ptr::addr_of!(ext.SubFormat)) };
        if subformat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT {
            format = SampleFormat::F32;
        } else {
            format = SampleFormat::I16;
        }
    }
    (rate, channels, format)
}

fn read_samples(ptr: *mut u8, frames: u32, channels: u16, format: SampleFormat) -> Vec<f32> {
    let sample_count = frames as usize * channels as usize;
    if ptr.is_null() || sample_count == 0 {
        return Vec::new();
    }
    unsafe {
        match format {
            SampleFormat::F32 => {
                let slice = std::slice::from_raw_parts(ptr as *const f32, sample_count);
                slice.to_vec()
            }
            SampleFormat::I16 => {
                let slice = std::slice::from_raw_parts(ptr as *const i16, sample_count);
                slice.iter().map(|s| *s as f32 / i16::MAX as f32).collect()
            }
        }
    }
}

fn update_peak(shared: &SharedState, kind: &SourceKind, samples: &[f32]) {
    let mut peak = 0f32;
    for sample in samples {
        let abs = sample.abs();
        if abs > peak {
            peak = abs;
        }
    }
    let value = (peak * i16::MAX as f32) as u32;
    match kind {
        SourceKind::Microphone => {
            shared.mic_peak.store(value, Ordering::Relaxed);
        }
        SourceKind::System => {
            shared.system_peak.store(value, Ordering::Relaxed);
        }
    }
}

fn to_stereo(samples: &[f32], channels: usize) -> Vec<f32> {
    if channels == TARGET_CHANNELS as usize {
        return samples.to_vec();
    }
    let frames = samples.len() / channels;
    let mut out = Vec::with_capacity(frames * TARGET_CHANNELS as usize);
    for frame in 0..frames {
        let base = frame * channels;
        let left = samples[base];
        let right = if channels > 1 {
            samples[base + 1]
        } else {
            left
        };
        out.push(left);
        out.push(right);
    }
    out
}

struct LinearResampler {
    input_rate: u32,
    output_rate: u32,
    channels: usize,
    pos: f64,
    buffer: Vec<f32>,
}

impl LinearResampler {
    fn new(input_rate: u32, output_rate: u32, channels: usize) -> Self {
        LinearResampler {
            input_rate,
            output_rate,
            channels,
            pos: 0.0,
            buffer: Vec::new(),
        }
    }

    fn push(&mut self, samples: &[f32]) -> Vec<f32> {
        self.buffer.extend_from_slice(samples);
        if self.input_rate == 0 || self.output_rate == 0 || self.channels == 0 {
            return Vec::new();
        }
        let step = self.input_rate as f64 / self.output_rate as f64;
        let frames_available = self.buffer.len() / self.channels;
        let mut out = Vec::new();
        while self.pos + 1.0 < frames_available as f64 {
            let i0 = self.pos.floor() as usize;
            let i1 = i0 + 1;
            let frac = self.pos - i0 as f64;
            for ch in 0..self.channels {
                let s0 = self.buffer[i0 * self.channels + ch];
                let s1 = self.buffer[i1 * self.channels + ch];
                out.push((1.0 - frac as f32) * s0 + (frac as f32) * s1);
            }
            self.pos += step;
        }
        let drop_frames = self.pos.floor() as usize;
        if drop_frames > 0 {
            let drop_samples = drop_frames * self.channels;
            self.buffer.drain(0..drop_samples);
            self.pos -= drop_frames as f64;
        }
        out
    }
}

struct WavWriter {
    file: File,
    data_size: u32,
}

impl WavWriter {
    fn create(path: &Path) -> Result<Self, String> {
        let file = OpenOptions::new()
            .create(true)
            .truncate(true)
            .write(true)
            .open(path)
            .map_err(|e| e.to_string())?;
        let mut writer = WavWriter { file, data_size: 0 };
        writer.write_header_placeholder()?;
        Ok(writer)
    }

    fn write_header_placeholder(&mut self) -> Result<(), String> {
        self.file.write_all(b"RIFF").map_err(|e| e.to_string())?;
        self.file
            .write_all(&0u32.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file.write_all(b"WAVE").map_err(|e| e.to_string())?;
        self.file.write_all(b"fmt ").map_err(|e| e.to_string())?;
        self.file
            .write_all(&16u32.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&1u16.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&TARGET_CHANNELS.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&TARGET_SAMPLE_RATE.to_le_bytes())
            .map_err(|e| e.to_string())?;
        let byte_rate = TARGET_SAMPLE_RATE * TARGET_CHANNELS as u32 * (TARGET_BITS as u32 / 8);
        let block_align = TARGET_CHANNELS * (TARGET_BITS / 8);
        self.file
            .write_all(&byte_rate.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&block_align.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&TARGET_BITS.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file.write_all(b"data").map_err(|e| e.to_string())?;
        self.file
            .write_all(&0u32.to_le_bytes())
            .map_err(|e| e.to_string())?;
        Ok(())
    }

    fn write_f32(&mut self, samples: &[f32]) -> Result<(), String> {
        let mut buf = Vec::with_capacity(samples.len() * 2);
        for sample in samples {
            let clamped = sample.clamp(-1.0, 1.0);
            let v = (clamped * i16::MAX as f32) as i16;
            buf.extend_from_slice(&v.to_le_bytes());
        }
        self.file.write_all(&buf).map_err(|e| e.to_string())?;
        self.data_size = self.data_size.saturating_add(buf.len() as u32);
        Ok(())
    }

    fn finalize(&mut self) -> Result<(), String> {
        let riff_size = 36u32.saturating_add(self.data_size);
        self.file
            .seek(SeekFrom::Start(4))
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&riff_size.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file
            .seek(SeekFrom::Start(40))
            .map_err(|e| e.to_string())?;
        self.file
            .write_all(&self.data_size.to_le_bytes())
            .map_err(|e| e.to_string())?;
        self.file.flush().map_err(|e| e.to_string())?;
        Ok(())
    }
}

pub fn default_output_folder() -> PathBuf {
    PathBuf::from(settings::default_podcast_save_folder())
}
