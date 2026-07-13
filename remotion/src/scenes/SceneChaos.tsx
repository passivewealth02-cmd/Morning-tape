import React from 'react'
import { AbsoluteFill, interpolate, useCurrentFrame } from 'remotion'
import { BRAND, FONT } from '../brand'
import { Chip } from '../ui'

const chips = [
  { label: 'Calls', x: -420, y: -170, d: 0 },
  { label: 'Texts', x: 380, y: -130, d: 4 },
  { label: 'Email', x: -320, y: 180, d: 8 },
  { label: 'Spreadsheets', x: 340, y: 200, d: 12 },
  { label: 'WhatsApp', x: 20, y: 270, d: 16 },
  { label: 'Voicemail', x: -460, y: 40, d: 20 },
]

export const SceneChaos: React.FC = () => {
  const frame = useCurrentFrame()
  const title = interpolate(frame, [8, 26], [0, 1], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })
  return (
    <AbsoluteFill style={{ background: BRAND.ink, justifyContent: 'center', alignItems: 'center' }}>
      {chips.map((c, i) => {
        const o = interpolate(frame, [c.d, c.d + 12], [0, 1], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })
        const drift = Math.sin((frame + i * 20) / 30) * 8
        return (
          <div
            key={c.label}
            style={{ position: 'absolute', transform: `translate(${c.x}px, ${c.y + drift}px)`, opacity: o * 0.9 }}
          >
            <Chip label={c.label} bg="rgba(255,255,255,0.08)" color="rgba(255,255,255,0.75)" />
          </div>
        )
      })}
      <div
        style={{
          fontFamily: FONT,
          fontWeight: 600,
          fontSize: 66,
          letterSpacing: -2,
          color: '#fff',
          textAlign: 'center',
          opacity: title,
          maxWidth: 1100,
        }}
      >
        Maintenance requests are chaos.
      </div>
    </AbsoluteFill>
  )
}
