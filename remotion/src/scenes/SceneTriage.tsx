import React from 'react'
import { AbsoluteFill, spring, useCurrentFrame, useVideoConfig } from 'remotion'
import { BRAND, FONT } from '../brand'
import { Caption } from '../ui'

const labels = [
  { text: 'Plumbing', bg: BRAND.indigo50, color: BRAND.indigo, at: 10 },
  { text: 'Emergency', bg: '#FEE2E2', color: BRAND.red, at: 18 },
  { text: 'Vendor: Plumber', bg: BRAND.gray100, color: BRAND.gray600, at: 26 },
]

export const SceneTriage: React.FC = () => {
  const frame = useCurrentFrame()
  const { fps } = useVideoConfig()
  const card = spring({ frame, fps, config: { damping: 16 }, durationInFrames: 22 })

  return (
    <AbsoluteFill style={{ background: BRAND.gray50, justifyContent: 'center', alignItems: 'center', gap: 56 }}>
      <div
        style={{
          transform: `scale(${card})`,
          width: 780,
          background: '#fff',
          borderRadius: 24,
          border: `1px solid ${BRAND.gray200}`,
          padding: 40,
          boxShadow: '0 24px 70px rgba(17,24,39,0.10)',
        }}
      >
        <div style={{ fontFamily: FONT, fontWeight: 700, fontSize: 34, color: BRAND.ink }}>
          Water leaking under kitchen sink
        </div>
        <div style={{ fontFamily: FONT, fontSize: 24, color: BRAND.gray500, marginTop: 8 }}>
          AI summary: Active leak under the sink — dispatch a plumber promptly.
        </div>
        <div style={{ display: 'flex', gap: 12, marginTop: 26 }}>
          {labels.map((l) => {
            const s = spring({ frame: frame - l.at, fps, config: { damping: 12 }, durationInFrames: 16 })
            return (
              <span
                key={l.text}
                style={{
                  transform: `scale(${s})`,
                  fontFamily: FONT,
                  fontWeight: 600,
                  fontSize: 28,
                  padding: '10px 22px',
                  borderRadius: 999,
                  background: l.bg,
                  color: l.color,
                }}
              >
                {l.text}
              </span>
            )
          })}
        </div>
      </div>
      <Caption delay={8}>AI triages it in seconds — trade, urgency, vendor.</Caption>
    </AbsoluteFill>
  )
}
