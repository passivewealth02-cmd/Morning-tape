import React from 'react'
import { Composition } from 'remotion'
import { Promo } from './Promo'
import { CinematicPromo, CinematicFramed, TOTAL_CINEMATIC_FRAMES } from './CinematicPromo'

const FPS = 30
const DURATION = 600 // 20 seconds

export const RemotionRoot: React.FC = () => {
  return (
    <>
      <Composition
        id="Cinematic"
        component={CinematicPromo}
        durationInFrames={TOTAL_CINEMATIC_FRAMES}
        fps={24}
        width={1920}
        height={1080}
      />
      <Composition
        id="CinematicVertical"
        component={CinematicFramed}
        durationInFrames={TOTAL_CINEMATIC_FRAMES}
        fps={24}
        width={1080}
        height={1920}
      />
      <Composition
        id="CinematicSquare"
        component={CinematicFramed}
        durationInFrames={TOTAL_CINEMATIC_FRAMES}
        fps={24}
        width={1080}
        height={1080}
      />
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
