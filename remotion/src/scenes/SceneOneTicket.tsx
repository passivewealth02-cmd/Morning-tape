import React from 'react'
import { AbsoluteFill, spring, useCurrentFrame, useVideoConfig } from 'remotion'
import { BRAND, FONT } from '../brand'
import { Caption, Chip } from '../ui'

export const SceneOneTicket: React.FC = () => {
  const frame = useCurrentFrame()
  const { fps } = useVideoConfig()
  const s = spring({ frame, fps, config: { damping: 16 }, durationInFrames: 26 })
  return (
    <AbsoluteFill style={{ background: BRAND.gray50, justifyContent: 'center', alignItems: 'center', gap: 56 }}>
      <div
        style={{
          transform: `scale(${s})`,
          width: 760,
          background: '#fff',
          borderRadius: 24,
          border: `1px solid ${BRAND.gray200}`,
          padding: 40,
          boxShadow: '0 24px 70px rgba(17,24,39,0.10)',
        }}
      >
        <div style={{ fontFamily: FONT, fontWeight: 700, fontSize: 36, color: BRAND.ink }}>
          Water leaking under kitchen sink
        </div>
        <div style={{ fontFamily: FONT, fontSize: 27, color: BRAND.gray500, marginTop: 10 }}>
          Maple Gardens · Unit 4B
        </div>
        <div style={{ display: 'flex', gap: 12, marginTop: 24 }}>
          <Chip label="New" bg={BRAND.indigo50} color={BRAND.indigo} />
          <Chip label="Just now" />
        </div>
      </div>
      <Caption delay={12}>One clean ticket. Nothing slips through.</Caption>
    </AbsoluteFill>
  )
}
