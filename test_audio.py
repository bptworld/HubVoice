#!/usr/bin/env python
"""Quick audio test for FPH Satellite-1 via ESPHome API."""
import asyncio
import aioesphomeapi

HOST = "192.168.4.86"
PORT = 6054

async def main():
    client = aioesphomeapi.APIClient(HOST, PORT, password=None)
    await client.connect(login=True)

    entities, _ = await client.list_entities_services()
    mp = next((e for e in entities if isinstance(e, aioesphomeapi.MediaPlayerInfo)), None)
    if not mp:
        print("No media player found!")
        await client.disconnect()
        return

    print(f"Found media player: {mp.name} (key={mp.key})")

    # Play a short test tone via URL
    test_url = "http://www.kozco.com/tech/piano2.wav"
    print(f"Playing test audio: {test_url}")
    client.media_player_command(
        mp.key,
        media_url=test_url,
        announcement=True,
    )
    print("Command sent. Listen for audio...")
    await asyncio.sleep(5)
    await client.disconnect()

asyncio.run(main())
