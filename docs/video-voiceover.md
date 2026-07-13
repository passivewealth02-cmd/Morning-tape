# Maintena Promo — Voiceover Script (20s)

Timed to the 7 scenes (30fps). Read at a calm, confident pace (~2.5 words/sec).
Total ≈ 55 words. Works for a human recording or an AI voice (ElevenLabs, etc.).

| Time | Scene | Voiceover | On-screen caption |
|---|---|---|---|
| 0:00–0:03 | Chaos | "Maintenance requests get lost in calls, texts, and email." | Maintenance requests are chaos. |
| 0:03–0:06 | One ticket | "Maintena brings every request into one place." | One clean ticket. Nothing slips through. |
| 0:06–0:09 | Scan | "Tenants just scan a QR code — no app, no login." | Tenants scan to report — the form knows their unit. |
| 0:09–0:12 | Triage | "AI reads it and sets the trade, urgency, and vendor — in seconds." | AI triages it in seconds. |
| 0:12–0:15 | Dispatch | "Then dispatches the right vendor automatically." | The right vendor, dispatched automatically. |
| 0:15–0:18 | Tracking | "And everyone stays in the loop, start to finish." | Everyone stays in the loop. |
| 0:18–0:20 | Logo | "Maintena — the AI operations layer for property maintenance." | trymaintena.com |

## Alt: punchier 15-second cut
"Maintenance is chaos. Maintena fixes it. Tenants scan a QR code, AI dispatches the right
vendor, and everyone stays in the loop. Maintena — property maintenance on autopilot."

## Adding it to the video
Drop your recorded/generated audio file into `remotion/public/vo.mp3`, then add to
`src/Promo.tsx`:
```tsx
import { Audio, staticFile } from 'remotion'
// inside <AbsoluteFill>, above <Series>:
<Audio src={staticFile('vo.mp3')} />
```
Background music works the same way (add a second `<Audio>` with `volume={0.15}`).
