import React from 'react'
import { AbsoluteFill, interpolate, spring, useCurrentFrame, useVideoConfig } from 'remotion'
import { BRAND, FONT } from '../brand'
import { Caption } from '../ui'

export const SceneDispatch: React.FC = () => {
  const frame = useCurrentFrame()
  const { fps } = useVideoConfig()
  const slide = spring({ frame, fps, config: { damping: 18 }, durationInFrames: 26 })
  const x = interpolate(slide, [0, 1], [420, 0])
  const check = spring({ frame: frame - 18, fps, config: { damping: 10 }, durationInFrames: 16 })

  return (
    <AbsoluteFill style={{ background: BRAND.white, justifyContent: 'center', alignItems: 'center', gap: 56 }}>
      <div
        style={{
          transform: `translateX(${x}px)`,
          opacity: slide,
          width: 720,
          background: '#fff',
          borderRadius: 24,
          border: `1px solid ${BRAND.gray200}`,
          padding: 36,
          boxShadow: '0 24px 70px rgba(17,24,39,0.10)',
          display: 'flex',
          alignItems: 'center',
          gap: 24,
        }}
      >
        <div
          style={{
            width: 84,
            height: 84,
            borderRadius: 20,
            background: BRAND.indigo50,
            color: BRAND.indigo,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontFamily: FONT,
            fontWeight: 700,
            fontSize: 34,
          }}
        >
          MP
        </div>
        <div style={{ flex: 1 }}>
          <div style={{ fontFamily: FONT, fontWeight: 700, fontSize: 32, color: BRAND.ink }}>Mike&apos;s Plumbing Co.</div>
          <div style={{ fontFamily: FONT, fontSize: 24, color: BRAND.gray500, marginTop: 4 }}>
            Plumber · ★ 4.8 · Available
          </div>
        </div>
        <span
          style={{
            transform: `scale(${check})`,
            fontFamily: FONT,
            fontWeight: 700,
            fontSize: 26,
            padding: '10px 22px',
            borderRadius: 999,
            background: '#DCFCE7',
            color: BRAND.green,
          }}
        >
          Assigned ✓
        </span>
      </div>
      <Caption delay={10}>The right vendor, dispatched automatically.</Caption>
    </AbsoluteFill>
  )
}
