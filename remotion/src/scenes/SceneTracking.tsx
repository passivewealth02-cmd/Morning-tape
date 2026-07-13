import React from 'react'
import { AbsoluteFill, interpolate, useCurrentFrame } from 'remotion'
import { BRAND, FONT } from '../brand'
import { Caption } from '../ui'

const steps = ['Received', 'Assigned', 'In progress', 'Completed']

export const SceneTracking: React.FC = () => {
  const frame = useCurrentFrame()
  // progress advances through the 4 steps over ~60 frames
  const progress = interpolate(frame, [10, 70], [0, 3], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })

  return (
    <AbsoluteFill style={{ background: BRAND.gray50, justifyContent: 'center', alignItems: 'center', gap: 64 }}>
      <div style={{ display: 'flex', alignItems: 'center' }}>
        {steps.map((s, i) => {
          const done = progress >= i
          const active = Math.floor(progress) === i
          return (
            <React.Fragment key={s}>
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', width: 200 }}>
                <div
                  style={{
                    width: 64,
                    height: 64,
                    borderRadius: 999,
                    background: done ? BRAND.green : active ? BRAND.indigo : BRAND.gray200,
                    color: '#fff',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    fontFamily: FONT,
                    fontWeight: 700,
                    fontSize: 28,
                    transition: 'background 0.2s',
                  }}
                >
                  {done ? '✓' : i + 1}
                </div>
                <div
                  style={{
                    marginTop: 14,
                    fontFamily: FONT,
                    fontWeight: 600,
                    fontSize: 24,
                    color: done || active ? BRAND.ink : BRAND.gray400,
                  }}
                >
                  {s}
                </div>
              </div>
              {i < steps.length - 1 && (
                <div style={{ width: 90, height: 4, background: BRAND.gray200, marginTop: -30, position: 'relative' }}>
                  <div
                    style={{
                      position: 'absolute',
                      inset: 0,
                      width: `${interpolate(progress - i, [0, 1], [0, 100], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })}%`,
                      background: BRAND.green,
                    }}
                  />
                </div>
              )}
            </React.Fragment>
          )
        })}
      </div>
      <Caption delay={8}>Everyone stays in the loop — automatically.</Caption>
    </AbsoluteFill>
  )
}
