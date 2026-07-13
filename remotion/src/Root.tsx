import React from 'react'
import { Composition } from 'remotion'
import { Promo } from './Promo'

const FPS = 30
const DURATION = 600 // 20 seconds

export const RemotionRoot: React.FC = () => {
  return (
    <>
      <Composition
        id="PromoLandscape"
        component={Promo}
        durationInFrames={DURATION}
        fps={FPS}
        width={1920}
        height={1080}
      />
      <Composition
        id="PromoVertical"
        component={Promo}
        durationInFrames={DURATION}
        fps={FPS}
        width={1080}
        height={1920}
      />
      <Composition
        id="PromoSquare"
        component={Promo}
        durationInFrames={DURATION}
        fps={FPS}
        width={1080}
        height={1080}
      />
    </>
  )
}
