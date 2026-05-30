# Bot financiero (Telegram)

Proyecto **independiente** en esta carpeta. No interfiere con la app móvil del repositorio.

## Requisitos

- Python 3.10 o superior
- (Opcional pero recomendado) Cuenta OpenAI y API key para transcribir voz con Whisper
- (Solo si no usas OpenAI) [ffmpeg](https://ffmpeg.org/) en el PATH, para convertir `.ogg` de Telegram antes de Google SpeechRecognition

## Instalación

Desde la raíz del repositorio (`proyecto-tesis-app-version-3.1`):

```bash
cd bot_financiero_telegram
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy env.example .env
```

Edita `.env` con tu token y tu ID de usuario (ver abajo).

## Cómo obtener el token de Telegram (BotFather)

1. Abre Telegram y busca el usuario **@BotFather**.
2. Inicia el chat y envía el comando `/newbot`.
3. Elige un **nombre** para el bot (visible al público) y un **username** que termine en `bot` (por ejemplo `mi_finanzas_local_bot`).
4. BotFather responderá con un mensaje que incluye el **HTTP API token** (larga cadena tipo `123456789:ABCdef...`). Ese valor va en `TELEGRAM_BOT_TOKEN` dentro de `.env`.
5. **No compartas** el token ni lo subas a Git (`.env` está en `.gitignore`).

### Tu ID de usuario (solo tú puedes usar el bot)

- Escribe a **@userinfobot** o **@getidsbot** en Telegram; te mostrarán tu **Id** numérico.
- Ese número va en `TELEGRAM_ALLOWED_USER_ID` en `.env`.

## Ejecutar el bot

Desde la **raíz del repositorio** (carpeta padre de `bot_financiero_telegram`):

```bash
python -m bot_financiero_telegram
```

Con el entorno virtual activado y dependencias instaladas.

## Uso

- **Texto o voz**: el bot reconoce ordenes y puede registrar retenciones en `consolidado_financiero.xlsx`.
- **Registro de retencion** (texto o audio transcrito) con formato recomendado:
  - `registrar retencion, fecha:24/04/2026, comprobante:20260400000381, rif:J-41278020-4, fecha factura:23/04/2026, nro factura:00007553, control:Z7C7018762, total:1.400,00, base imponible:1.206,90, iva retenido:144,83`
- **Reportes**: `retenciones recibidas del 01/05/2026 al 31/05/2026 en excel` o `... en pdf`.
- Cualquier otro usuario recibirá «Acceso no autorizado.»

## Notas

- El flujo de imagen/canal esta desactivado; la carga de datos es por texto o nota de voz.
- Si todo falla con «Acceso no autorizado», envia el comando `/mi_id` al bot y revisa que `TELEGRAM_ALLOWED_USER_ID` en `.env` sea exactamente ese numero (luego reinicia el proceso).
- Sin `OPENAI_API_KEY`, la voz usa Google + **ffmpeg**; si falla la nota de voz, instala ffmpeg en el PATH o configura Whisper.
