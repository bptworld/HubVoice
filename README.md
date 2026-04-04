# HubVoiceSat

This is the clean restart workspace.

Files:
- `home-assistant-voice-official.yaml`: untouched official HA Voice PE YAML from `esphome/home-assistant-voice-pe`
- `hubvoice-sat.yaml`: working copy we will adapt from that official baseline
- `secrets.yaml`: local-only development file; keep `wifi_ssid` and `wifi_password` empty before creating distributable release bins
- `build-hubvoice-sat.ps1`: repeatable local config/compile entrypoint for this workspace
- `build-setup-launcher.ps1`: publishes the Windows setup launcher exe and refreshes the repo-root exe
- `flash.bat`: simple Windows wrapper for config / compile / flash commands
- `HubVoiceSatSetup.exe`: single Windows launcher exe that self-stages the setup payload and opens the local setup page
- `.envs\`: local virtual environments (`main`, `runtime`, `airplay`) used by build/runtime scripts
- `setup-web.ps1`: setup page/service payload embedded into the launcher
- `build\HubVoiceSatSetup\HubVoiceSatSetup.exe`: published Windows launcher exe
- `patch.exe` / `patch.cmd` / `patch-wrapper.ps1`: local Windows patch shim so `micro-opus` can build without Git installed

Intent:
- start from the official HA Voice PE wake/button lifecycle
- verify stock wake word and stock button behavior first
- only then re-add Hubitat, custom UI, and helper/server pieces one layer at a time

Rule for this workspace:
- keep `home-assistant-voice-official.yaml` unchanged
- make all edits in `hubvoice-sat.yaml`

Current state:
- `hubvoice-sat.yaml` validates cleanly
- full `esphome compile hubvoice-sat.yaml` succeeds in this workspace
- stock/new satellites should be flashed once by USB first
- after HubVoiceSat firmware is on the device, future updates can be done by OTA from the satellite web page
- the firmware is now built as one generic image and uses `name_add_mac_suffix: true` so each unit gets a unique network identity
- end users can set a per-device local label from the device web page using `Satellite Name`; `Effective Satellite Name` shows the current result

Recommended end-user update flow:
- first install on a stock/unknown satellite: USB
- later updates on a satellite already running HubVoiceSat: OTA package
- use `.\build-ota-release.ps1` to generate the OTA handoff zip for end users
- `build-ota-release.ps1` will fail if `secrets.yaml` has non-empty Wi-Fi values to prevent embedding personal credentials in shipped firmware
- `build-ota-release.ps1` also scans the compiled `.bin` files and fails if local Wi-Fi, satellite, Hubitat, callback, or setup values from this workspace are found inside them

Suggested first-time USB path for end users:
- plug the satellite into USB
- open `https://web.esphome.io/`
- connect to the device
- choose the generated `*-factory.bin`
- after that first install, use the `*-ota.bin` from the satellite web page for later updates

Useful commands:
- `.\build-hubvoice-sat.ps1 -Action config`
- `.\build-hubvoice-sat.ps1 -Action compile`
- `.\build-setup-launcher.ps1`
- `.\build-ota-release.ps1`
- `.\verify-firmware-bins.ps1 -BinPaths .\hubvoice-sat-2026.03.31.6-factory.bin,.\hubvoice-sat-2026.03.31.6-ota.bin`
- `.\HubVoiceSatSetup.exe`
- `flash.bat config`
- `flash.bat compile`
- `flash.bat usb`
- `flash.bat ota`

Hubitat satellite TTS proxy:
- `HubVoiceSatelliteTTSDriver.groovy` is a custom Hubitat virtual driver that exposes a satellite as a `SpeechSynthesis` and `Notification` device.
- Create one virtual device per satellite in Hubitat using that driver.
- Set `runtimeBaseUrl` to the HubVoice runtime host, for example `http://192.168.1.50:8080`.
- Set `satelliteId` to the exact satellite ID from `satellites.csv`.
- Use the driver's `testConnection` command to verify the runtime is reachable.
- Use the driver's `discoverSatellites` command to pull the configured satellite IDs from the runtime.
- Use the driver's `testMessage` command or any Hubitat app `speak()` / notification action to send speech to that satellite.
- The runtime now exposes `/satellites` and returns the configured IDs/hosts loaded from `satellites.csv`.

Local timers and alarms:
- The runtime now handles timers and alarms directly before falling back to Hubitat.
- Timer phrases supported: `set a timer for 5 minutes`, `start a timer for one hour`, `cancel timer`, `cancel all timers`, `how much time is left on the timer`.
- Alarm phrases supported: `set an alarm for 7:30 am`, `set an alarm for 6 tomorrow morning`, `wake me up at 8 pm`, `cancel alarm`, `cancel all alarms`, `when is my next alarm`.
- Timers emit native ESPHome timer events so the existing LED progress and timer-finished ringing behavior on the satellite still works.
- Alarms are scheduled locally in the runtime and use the same satellite ringing path when they fire.
- Fallback ringing now repeats until dismissed, with an automatic stop after 1 minute.
- Dismiss phrases: `dismiss`, `silence`, `stop ringing`, `dismiss alarm`, `dismiss timer`, `stop alarm`, `stop timer`.
- Scheduler state is visible from the runtime root page, `/health`, and `/schedules`.
