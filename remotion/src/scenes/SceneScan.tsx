import React from 'react'
import { AbsoluteFill, interpolate, spring, useCurrentFrame, useVideoConfig } from 'remotion'
import { BRAND, FONT } from '../brand'
import { Caption, Chip } from '../ui'

// A simple stylized QR block built from a fixed pattern.
const QR_PATTERN = [
  '1111111 0 1011 0 1111111',
  '1000001 0 0110 0 1000001',
  '1011101 0 1101 0 1011101',
  '1011101 0 0011 0 1011101',
  '1011101 0 1110 0 1011101',
  '1000001 0 0101 0 1000001',
  '1111111 0 1010 0 1111111',
  '0000000 0 0000 0 0000000',
  '1101011 1 0110 1 0101101',
  '0010110 0 1011 0 1100010',
  '1100101 1 0101 1 0011011',
]

export const SceneScan: React.FC = () => {
  const frame = useCurrentFrame()
  const { fps } = useVideoConfig()
  const s = spring({ frame, fps, config: { damping: 16 }, durationInFrames: 24 })
  const scan = interpolate(frame % 60, [0, 60], [0, 1])
  const fill = interpolate(frame, [40, 55], [0, 1], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })

  const cell = 12
  return (
    <AbsoluteFill style={{ background: BRAND.white, justifyContent: 'center', alignItems: 'center', gap: 48 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 64, transform: `scale(${s})` }}>
        {/* Phone */}
        <div
          style={{
            width: 300,
            height: 600,
            borderRadius: 44,
            background: BRAND.ink,
            padding: 16,
            boxShadow: '0 30px 80px rgba(17,24,39,0.25)',
          }}
        >
          <div
            style={{
              width: '100%',
              height: '100%',
              borderRadius: 30,
              background: BRAND.gray50,
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              position: 'relative',
              overflow: 'hidden',
            }}
          >
            <div style={{ position: 'relative', background: '#fff', padding: 16, borderRadius: 12 }}>
              {QR_PATTERN.map((row, y) => (
                <div key={y} style={{ display: 'flex' }}>
                  {row.replace(/ /g, '').split('').map((v, x) => (
                    <div
                      key={x}
                      style={{ width: cell, height: cell, background: v === '1' ? BRAND.ink : 'transparent' }}
                    />
                  ))}
                </div>
              ))}
              {/* scan line */}
              <div
                style={{
                  position: 'absolute',
                  left: 8,
                  right: 8,
                  height: 4,
                  top: `${scan * 100}%`,
                  background: BRAND.indigo,
                  boxShadow: `0 0 16px ${BRAND.indigo}`,
                }}
              />
            </div>
            <div style={{ marginTop: 22, fontFamily: FONT, fontWeight: 600, fontSize: 20, color: BRAND.gray500 }}>
              Scan to report an issue
            </div>
          </div>
        </div>

        {/* Auto-filled location card */}
        <div style={{ opacity: fill, transform: `translateX(${(1 - fill) * -20}px)` }}>
          <div
            style={{
              background: BRAND.indigo50,
              border: `1px solid ${BRAND.indigo100}`,
              borderRadius: 18,
              padding: '20px 26px',
            }}
          >
            <div style={{ fontFamily: FONT, fontWeight: 600, fontSize: 22, color: BRAND.indigoDark }}>Reporting for</div>
            <div style={{ fontFamily: FONT, fontWeight: 700, fontSize: 30, color: BRAND.ink, marginTop: 4 }}>
              Maple Gardens · Unit 4B
            </div>
          </div>
          <div style={{ marginTop: 14 }}>
            <Chip label="No typing · No app · No login" bg={BRAND.gray100} color={BRAND.gray600} />
          </div>
        </div>
      </div>
      <Caption delay={10}>Tenants scan to report — the form knows their unit.</Caption>
    </AbsoluteFill>
  )
}
