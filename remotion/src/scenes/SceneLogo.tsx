import React from 'react'
import { AbsoluteFill, interpolate, useCurrentFrame } from 'remotion'
import { BRAND, FONT } from '../brand'
import { MaintenaLogo } from '../MaintenaLogo'

export const SceneLogo: React.FC = () => {
  const frame = useCurrentFrame()
  const tag = interpolate(frame, [30, 44], [0, 1], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })
  const url = interpolate(frame, [40, 52], [0, 1], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })

  return (
    <AbsoluteFill style={{ background: BRAND.white, justifyContent: 'center', alignItems: 'center' }}>
      <MaintenaLogo />
      <div
        style={{
          marginTop: 40,
          fontFamily: FONT,
          fontWeight: 600,
          fontSize: 40,
          letterSpacing: -1,
          color: BRAND.gray600,
          opacity: tag,
          textAlign: 'center',
        }}
      >
        The AI operations layer for property maintenance.
      </div>
      <div
        style={{
          marginTop: 20,
          fontFamily: FONT,
          fontWeight: 600,
          fontSize: 32,
          color: BRAND.indigo,
          opacity: url,
        }}
      >
        trymaintena.com
      </div>
    </AbsoluteFill>
  )
}
