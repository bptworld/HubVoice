# HubVoice and HubVoiceSat

Local-first voice control for Hubitat with ESPHome-based satellites, Windows runtime tools, and packaged release workflows.

This project combines:
- HubVoice Hubitat app logic
- HubVoiceSat firmware variants for supported satellite hardware
- HubVoice runtime services and control UI
- Build and release scripts for USB-first install and OTA updates

## What This Project Does

HubVoice provides natural voice control and status for your home, while HubVoiceSat satellites handle wake word, audio playback, and local device interaction.

Key capabilities:
- Local voice assistant flow with Hubitat integration
- Device control, status queries, and group commands
- Runtime scheduling for timers and alarms
- Satellite control deck for media, volume, and playback actions
- Optional AI fallback for non-control questions (Gemini or ChatGPT)
- Repeatable release packaging for factory and OTA firmware

## Important Notes

- First install on a device should be done with a factory firmware image over USB.
- Updates after first install should use OTA firmware from the satellite web UI.
- Release packaging checks for accidental secret leakage in binaries.
- Keep local Wi-Fi credentials out of distributable release artifacts.

## Firmware Variants

Common variants in this repo:
- hubvoice-sat.yaml: HA Voice PE default variant
- hubvoice-sat-fph.yaml: FutureProofHomes Satellite-1 variant
- hubvoice-sat-fph-ld2410.yaml: Satellite-1 + LD2410 presence variant
- hubvoice-sat-echos3r.yaml: EchoS3R-targeted variant

The OTA release script can package multiple variants into one release folder.

## Quick Start (Windows)

1. Clone this repository and open PowerShell in the repo root.
2. Configure and compile firmware:

```powershell
.\build-hubvoice-sat.ps1 -Action config
.\build-hubvoice-sat.ps1 -Action compile
```

3. For a first-time satellite flash, use a factory image via USB:
- Open https://web.esphome.io/
- Connect device
- Install matching factory bin

4. For ongoing updates, generate and use OTA release assets:

```powershell
.\build-ota-release.ps1
```

5. Launch the Windows setup runtime UI if needed:

```powershell
.\HubVoiceSatSetup.exe
```

## Build and Release Commands

Core build commands:

```powershell
.\build-hubvoice-sat.ps1 -Action config
.\build-hubvoice-sat.ps1 -Action compile
.\build-runtime-exe.ps1
.\build-setup-launcher.ps1
.\build-ota-release.ps1
```

Helpful wrappers:

```powershell
flash.bat config
flash.bat compile
flash.bat usb
flash.bat ota
```

Binary verification example:

```powershell
.\verify-firmware-bins.ps1 -BinPaths .\hubvoice-sat-2026.04.05.1-factory.bin,.\hubvoice-sat-2026.04.05.1-ota.bin
```

## Hubitat Integration

Main app files:
- HubVoice.groovy
- HubVoice_Controller.groovy
- HubVoice_Satellite_TTS.groovy

Satellite TTS proxy usage:
- Create a virtual device per satellite in Hubitat
- Point runtime base URL to your HubVoice runtime host
- Set satellite ID to match satellites.csv
- Use speak and notification commands to target specific satellites

## AI Fallback

When HubVoice cannot answer a non-control question, optional AI fallback can be used.

Supported providers:
- Gemini
- ChatGPT

Current behavior includes:
- Provider selection in app settings
- Provider-specific API key and model fields
- Basic usage and cost estimation tracking in diagnostics
- Reset controls for tracked cost history

Note: pricing and cost summaries are estimates based on configured assumptions and recorded token usage.

## Local Runtime Features

Runtime highlights include:
- Satellite control API endpoints
- Media command handling and pooling
- Timer and alarm scheduling
- Health and diagnostics endpoints
- Control latency telemetry endpoints

Examples of supported voice intents include:
- Device on/off and status checks
- Group actions like all lights and lock flows
- Timer and alarm create, query, cancel, dismiss

## Repository Layout

Primary areas:
- Firmware YAML: hubvoice-sat*.yaml
- Hubitat app logic: *.groovy
- Runtime service: hubvoice-runtime.py
- Setup and release scripts: *.ps1, *.bat
- Release outputs: releases/
- Runtime and setup assets: assets/, build/, setup-launcher/

## Troubleshooting

If setup or flashing fails:
- Confirm correct firmware variant for your hardware
- Verify first install used factory image, not OTA image
- Re-check USB cable and browser permissions in ESPHome Web Tools

If runtime control is slow or unstable:
- Confirm satellite host/IP entries in satellites.csv
- Verify runtime is reachable from the control UI and Hubitat
- Check health endpoints and runtime logs for connection errors

If release verification fails:
- Remove personal credentials from secrets.yaml
- Rebuild firmware and rerun verification scripts

## Contributing

Contributions are welcome.

Recommended process:
1. Open an issue describing bug or feature request.
2. Keep changes focused and testable.
3. Include clear notes for firmware/runtime impacts.
4. Validate build and release scripts before submitting.

## Credits

This project builds on:
- ESPHome
- Home Assistant voice satellite ecosystem
- Hubitat platform and community workflows
- FutureProofHomes Satellite hardware efforts

## License

See LICENSE in this repository for license terms.
