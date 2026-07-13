import React from 'react'
import { interpolate, spring, useCurrentFrame, useVideoConfig } from 'remotion'
import { BRAND, FONT } from './brand'

export const MaintenaLogo: React.FC<{ wordmark?: boolean; scale?: number }> = ({
  wordmark = true,
  scale = 1,
}) => {
  const frame = useCurrentFrame()
  const { fps } = useVideoConfig()

  const draw = interpolate(frame, [4, 26], [1, 0], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })
  const pop = spring({ frame, fps, config: { damping: 14 }, durationInFrames: 24 })
  const word = interpolate(frame, [22, 38], [0, 1], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })

  const tile = 160 * scale

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 24 * scale, transform: `scale(${pop})` }}>
      <div
        style={{
          width: tile,
          height: tile,
          borderRadius: 36 * scale,
          background: BRAND.indigo,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
        }}
      >
        <svg viewBox="0 0 120 120" width={tile / 2} height={tile / 2} fill="none">
          <path
            d="M 24 96 L 24 30 L 60 70 L 96 30 L 96 96"
            stroke="#fff"
            strokeWidth={16}
            strokeLinecap="round"
            strokeLinejoin="round"
            pathLength={1}
            strokeDasharray={1}
            strokeDashoffset={draw}
          />
        </svg>
      </div>
      {wordmark && (
        <span
          style={{
            fontFamily: FONT,
            fontWeight: 600,
            fontSize: 88 * scale,
            letterSpacing: -3 * scale,
            color: BRAND.ink,
            opacity: word,
            transform: `translateX(${(1 - word) * 20}px)`,
          }}
        >
          Maintena
        </span>
      )}
    </div>
  )
}
