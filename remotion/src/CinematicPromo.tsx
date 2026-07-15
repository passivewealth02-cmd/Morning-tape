import React from 'react'
import { AbsoluteFill, Series, useVideoConfig } from 'remotion'
import {
  CineColdOpen,
  CineQuestion,
  CineScan,
  CineTriage,
  CineDispatch,
  CineLoop,
  CineFinale,
} from './cinematic/scenes'

// 24fps for the filmic feel · 576 frames = 24 seconds
const T = {
  coldOpen: 96, // 4.0s
  question: 72, // 3.0s
  scan: 84, // 3.5s
  triage: 84, // 3.5s
  dispatch: 72, // 3.0s
  loop: 72, // 3.0s
  finale: 96, // 4.0s
}

export const TOTAL_CINEMATIC_FRAMES = Object.values(T).reduce((a, b) => a + b, 0)

export const CinematicPromo: React.FC = () => {
  return (
    <AbsoluteFill style={{ backgroundColor: '#000' }}>
      <Series>
        <Series.Sequence durationInFrames={T.coldOpen}>
          <CineColdOpen duration={T.coldOpen} />
        </Series.Sequence>
        <Series.Sequence durationInFrames={T.question}>
          <CineQuestion duration={T.question} />
        </Series.Sequence>
        <Series.Sequence durationInFrames={T.scan}>
          <CineScan duration={T.scan} />
        </Series.Sequence>
        <Series.Sequence durationInFrames={T.triage}>
          <CineTriage duration={T.triage} />
        </Series.Sequence>
        <Series.Sequence durationInFrames={T.dispatch}>
          <CineDispatch duration={T.dispatch} />
        </Series.Sequence>
        <Series.Sequence durationInFrames={T.loop}>
          <CineLoop duration={T.loop} />
        </Series.Sequence>
        <Series.Sequence durationInFrames={T.finale}>
          <CineFinale duration={T.finale} />
        </Series.Sequence>
      </Series>
    </AbsoluteFill>
  )
}

// Fits the 1920x1080 film onto any frame size (vertical / square) by scaling
// it to the target width and centering it on the dark backdrop — nothing
// overflows and every scene stays intact.
export const CinematicFramed: React.FC = () => {
  const { width } = useVideoConfig()
  const scale = width / 1920
  return (
    <AbsoluteFill style={{ background: '#07090F', justifyContent: 'center', alignItems: 'center' }}>
      <div style={{ position: 'relative', width: 1920, height: 1080, transform: `scale(${scale})` }}>
        <CinematicPromo />
      </div>
    </AbsoluteFill>
  )
}
