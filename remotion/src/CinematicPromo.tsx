import React from 'react'
import { AbsoluteFill, Series } from 'remotion'
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
