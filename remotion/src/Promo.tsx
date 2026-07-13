import React from 'react'
import { AbsoluteFill, Series } from 'remotion'
import { BRAND } from './brand'
import { SceneChaos } from './scenes/SceneChaos'
import { SceneOneTicket } from './scenes/SceneOneTicket'
import { SceneScan } from './scenes/SceneScan'
import { SceneTriage } from './scenes/SceneTriage'
import { SceneDispatch } from './scenes/SceneDispatch'
import { SceneTracking } from './scenes/SceneTracking'
import { SceneLogo } from './scenes/SceneLogo'

// 30fps · 600 frames = 20 seconds
export const Promo: React.FC = () => {
  return (
    <AbsoluteFill style={{ backgroundColor: BRAND.white }}>
      <Series>
        <Series.Sequence durationInFrames={90}>
          <SceneChaos />
        </Series.Sequence>
        <Series.Sequence durationInFrames={90}>
          <SceneOneTicket />
        </Series.Sequence>
        <Series.Sequence durationInFrames={90}>
          <SceneScan />
        </Series.Sequence>
        <Series.Sequence durationInFrames={90}>
          <SceneTriage />
        </Series.Sequence>
        <Series.Sequence durationInFrames={90}>
          <SceneDispatch />
        </Series.Sequence>
        <Series.Sequence durationInFrames={90}>
          <SceneTracking />
        </Series.Sequence>
        <Series.Sequence durationInFrames={60}>
          <SceneLogo />
        </Series.Sequence>
      </Series>
    </AbsoluteFill>
  )
}
