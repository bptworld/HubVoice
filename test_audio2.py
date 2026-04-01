import asyncio, aioesphomeapi
from datetime import datetime

async def main():
    client = aioesphomeapi.APIClient('192.168.4.86', 6054, password=None)
    await client.connect(login=True)

    all_states = []
    def on_state(s):
        if s.key == 2232357057:
            ts = datetime.now().strftime("%H:%M:%S.%f")
            all_states.append((ts, s.state, s.volume))
            print(f"  [{ts}] state={s.state} vol={s.volume}")

    client.subscribe_states(on_state)
    await asyncio.sleep(3)  # watch for self-triggered state changes

    print(f"Baseline states before command: {all_states}")
    print("Sending play command...")
    client.media_player_command(2232357057, media_url='http://192.168.4.23:8766/test_tone.wav', announcement=True)
    await asyncio.sleep(8)
    print("Done.")
    await client.disconnect()

asyncio.run(main())
