#!/usr/bin/env bash
# Renders the cinematic commercial in all three aspect ratios and lays Brian's
# ElevenLabs voiceover (audio/brian.mp3) onto each, synced to the scenes.
#
# Usage:  bash scripts/add-voiceover.sh
# Output: out/maintena-commercial-brian{,-vertical,-square}.mp4
set -euo pipefail
cd "$(dirname "$0")/.."

VO=audio/brian.mp3
mkdir -p out

# Brian's read (one take with pauses) split at the pauses and each line placed
# on its scene. Offsets in ms match the Cinematic composition scene starts.
FILTER="[1:a]atrim=0:2.95,asetpts=PTS-STARTPTS,adelay=400:all=1[a1];\
[1:a]atrim=3.42:5.00,asetpts=PTS-STARTPTS,adelay=4300:all=1[a2];\
[1:a]atrim=5.62:7.70,asetpts=PTS-STARTPTS,adelay=7400:all=1[a3];\
[1:a]atrim=8.18:11.62,asetpts=PTS-STARTPTS,adelay=10700:all=1[a4];\
[1:a]atrim=12.26:14.05,asetpts=PTS-STARTPTS,adelay=14300:all=1[a5];\
[1:a]atrim=14.33:16.65,asetpts=PTS-STARTPTS,adelay=17300:all=1[a6];\
[1:a]atrim=17.03:20.30,asetpts=PTS-STARTPTS,adelay=20400:all=1[a7];\
[a1][a2][a3][a4][a5][a6][a7]amix=inputs=7:normalize=0,apad,atrim=0:24,loudnorm=I=-16:TP=-1.5[vo]"

render_and_mux () {
  local comp="$1" silent="$2" final="$3"
  npx remotion render "$comp" "out/$silent"
  npx remotion ffmpeg -y -i "out/$silent" -i "$VO" \
    -filter_complex "$FILTER" -map 0:v -map "[vo]" -c:v copy -c:a aac -b:a 192k "out/$final"
  echo "  -> out/$final"
}

render_and_mux Cinematic          cine-landscape-silent.mp4 maintena-commercial-brian.mp4
render_and_mux CinematicVertical  cine-vertical-silent.mp4  maintena-commercial-brian-vertical.mp4
render_and_mux CinematicSquare    cine-square-silent.mp4    maintena-commercial-brian-square.mp4

echo "Done. Final videos are in out/."
