def listen_text_once():
    # Whisper-only: no SpeechRecognition, no PyAudio
    import sounddevice as sd
    import numpy as np
    from faster_whisper import WhisperModel

    # init model once and cache on function attribute
    model = getattr(listen_text_once, "_model", None)
    if model is None:
        # pick a small model to start: "base" or "tiny"
        listen_text_once._model = model = WhisperModel("base", device="cpu", compute_type="int8")

    fs = 16000
    duration = 4.0  # seconds of audio per turn
    print("Listening… (Whisper)")
    audio = sd.rec(int(duration * fs), samplerate=fs, channels=1, dtype="float32")
    sd.wait()

    # Whisper expects either a path or a 1-D float32 array
    mono = audio.squeeze().astype(np.float32)

    segments, _ = model.transcribe(mono, language="en")
    text = " ".join(s.text for s in segments).strip()
    return text
