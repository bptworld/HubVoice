# I2S Wake-Word Triage (Manual Hardware Bisect)

This workflow finds the first bad commit using actual device behavior.

## 1) Start from the debug branch snapshot

- Branch: debug/i2s-triage-2026-04-12
- Known rollback snapshot commit: 4cc16b6

## 2) Define pass/fail rule (strict)

A commit is:

- good: device boots, no repeating I2S microphone errors, wake word triggers within 60 seconds.
- bad: repeating errors like "Failed to start I2S channel" or wake word does not trigger in 60 seconds.

## 3) Run bisect

```powershell
git bisect start
git bisect bad
git bisect good 9a80cdd
```

Then for each checked out commit:

```powershell
& c:\HubVoice\.envs\runtime\Scripts\Activate.ps1
python -m esphome -s device_name livingroom2 -s friendly_name Livingroom2 run .\hubvoice-sat-fph.yaml --device 192.168.4.126 --no-logs
# Observe behavior for <=60s after boot
```

Mark result:

```powershell
# If wake word works and no I2S spam:
git bisect good

# If I2S errors or no wake response:
git bisect bad
```

Optional logging each step:

```powershell
.\triage\record-result.ps1 -Device livingroom2 -Outcome good -WakeWord worked -I2S repeating-errors -Notes "example"
```

## 4) Finish

```powershell
git bisect log > triage\bisect-session.log
git bisect reset
```

When bisect finishes, Git prints the first bad commit.
