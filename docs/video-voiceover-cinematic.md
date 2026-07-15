# Cinematic Commercial — Voiceover Script (24s)

Timed to the `Cinematic` Remotion composition (24fps, 576 frames). Deliver in a calm,
low, confident "movie trailer" read — unhurried, slight pauses at the periods.

| Start | Window | Line |
|---|---|---|
| 0:00.4 | 3.6s | Every day, maintenance requests slip through the cracks. |
| 0:04.3 | 2.7s | What if maintenance ran itself? |
| 0:07.4 | 3.1s | Tenants just scan the code on their door. |
| 0:10.7 | 3.3s | AI reads it, grades it, and routes it — in seconds. |
| 0:14.3 | 2.7s | The right pro is already on the way. |
| 0:17.3 | 2.7s | And no one has to ask for an update again. |
| 0:20.4 | 3.6s | Maintena. Property maintenance, on autopilot. |

Full paragraph (for pasting into ElevenLabs as one take):

> Every day, maintenance requests slip through the cracks. … What if maintenance ran
> itself? … Tenants just scan the code on their door. … AI reads it, grades it, and routes
> it — in seconds. … The right pro is already on the way. … And no one has to ask for an
> update again. … Maintena. Property maintenance, on autopilot.

ElevenLabs settings that work well: a deep narration voice (e.g. "Brian"/"Daniel" style),
Stability ~50%, Similarity ~75%, Style ~15%, speed 0.95.

## Replacing the placeholder voice with a real one

The current `out/maintena-cinematic-vo.mp4` uses an offline synthetic voice. To swap in a
premium track (single MP3 of the whole read):

```bash
cd remotion
npx remotion ffmpeg -y -i out/maintena-cinematic.mp4 -i YOUR-VO.mp3 \
  -filter_complex "[1:a]apad,atrim=0:24,loudnorm=I=-16:TP=-1.5[vo]" \
  -map 0:v -map "[vo]" -c:v copy -c:a aac -b:a 192k out/maintena-cinematic-final.mp4
```

If instead you export the 7 lines as separate files, the per-line offsets (ms) are:
400, 4300, 7400, 10700, 14300, 17300, 20400 — mix with `adelay` + `amix` as in the repo
history, or just ask Claude to do it.
