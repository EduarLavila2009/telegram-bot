"""Transcripción de notas de voz: OpenAI Whisper (preferido) o SpeechRecognition."""

from __future__ import annotations

import io
import tempfile
from pathlib import Path

from . import config


def transcribe_audio_file(path: Path) -> str:
    path = Path(path)
    if config.OPENAI_API_KEY:
        return _transcribe_openai(path)
    return _transcribe_google(path)


def _transcribe_openai(path: Path) -> str:
    from openai import OpenAI

    client = OpenAI(api_key=config.OPENAI_API_KEY)
    with path.open("rb") as audio_f:
        data = audio_f.read()
    bio = io.BytesIO(data)
    bio.name = path.name or "voice.ogg"
    result = client.audio.transcriptions.create(
        model="whisper-1",
        file=bio,
        language="es",
    )
    return (result.text or "").strip()


def _transcribe_google(path: Path) -> str:
    import speech_recognition as sr
    from pydub import AudioSegment

    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
        wav_path = Path(tmp.name)
    try:
        try:
            seg = AudioSegment.from_file(path)
        except Exception:
            seg = AudioSegment.from_file(path, format="ogg")
        seg.export(str(wav_path), format="wav")
        r = sr.Recognizer()
        with sr.AudioFile(str(wav_path)) as source:
            audio = r.record(source)
        return r.recognize_google(audio, language="es-VE")
    finally:
        if wav_path.exists():
            wav_path.unlink(missing_ok=True)
