import React from 'react'
import { AbsoluteFill, interpolate, useCurrentFrame, useVideoConfig } from 'remotion'

export const CINE = {
  bg: '#07090F', // near-black with a cold cast
  bg2: '#0B0E1A',
  glass: 'rgba(255,255,255,0.05)',
  glassBorder: 'rgba(255,255,255,0.10)',
  textDim: 'rgba(255,255,255,0.55)',
  textFaint: 'rgba(255,255,255,0.30)',
}

/** 2.39:1-style black bars for the film look. */
export const Letterbox: React.FC = () => {
  const { height } = useVideoConfig()
  const bar = Math.round(height * 0.1)
  return (
    <>
      <div style={{ position: 'absolute', top: 0, left: 0, right: 0, height: bar, background: '#000', zIndex: 40 }} />
      <div style={{ position: 'absolute', bottom: 0, left: 0, right: 0, height: bar, background: '#000', zIndex: 40 }} />
    </>
  )
}

/** Darkened corners. */
export const Vignette: React.FC = () => (
  <AbsoluteFill
    style={{
      background: 'radial-gradient(ellipse at center, rgba(0,0,0,0) 50%, rgba(0,0,0,0.6) 100%)',
      zIndex: 30,
      pointerEvents: 'none',
    }}
  />
)

/** A slow-drifting indigo light source behind the content. */
export const AmbientGlow: React.FC<{ x?: number; y?: number; hue?: string; size?: number }> = ({
  x = 50,
  y = 45,
  hue = '79,70,229',
  size = 60,
}) => {
  const frame = useCurrentFrame()
  const dx = Math.sin(frame / 90) * 4
  const dy = Math.cos(frame / 110) * 3
  return (
    <AbsoluteFill
      style={{
        background: `radial-gradient(circle at ${x + dx}% ${y + dy}%, rgba(${hue},0.22) 0%, rgba(${hue},0.07) ${size * 0.55}%, rgba(0,0,0,0) ${size}%)`,
        zIndex: 1,
      }}
    />
  )
}

/**
 * Scene shell: fades from/to black at the edges and applies a slow camera push-in.
 * Pass the scene's duration in frames.
 */
export const Scene: React.FC<{
  duration: number
  from?: number
  to?: number
  children: React.ReactNode
}> = ({ duration, from = 1.0, to = 1.06, children }) => {
  const frame = useCurrentFrame()
  const fadeIn = interpolate(frame, [0, 10], [0, 1], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })
  const fadeOut = interpolate(frame, [duration - 10, duration - 1], [1, 0], {
    extrapolateLeft: 'clamp',
    extrapolateRight: 'clamp',
  })
  const scale = interpolate(frame, [0, duration], [from, to])
  return (
    <AbsoluteFill style={{ background: CINE.bg }}>
      <AbsoluteFill
        style={{
          opacity: Math.min(fadeIn, fadeOut),
          transform: `scale(${scale})`,
          justifyContent: 'center',
          alignItems: 'center',
        }}
      >
        {children}
      </AbsoluteFill>
      <Vignette />
      <Letterbox />
    </AbsoluteFill>
  )
}

/** Word-by-word cinematic text reveal. */
export const RevealText: React.FC<{
  text: string
  delay?: number
  perWord?: number
  size?: number
  weight?: number
  color?: string
  font: string
  glow?: boolean
  tracking?: number
}> = ({ text, delay = 0, perWord = 4, size = 76, weight = 600, color = '#fff', font, glow = false, tracking = -2 }) => {
  const frame = useCurrentFrame()
  const words = text.split(' ')
  return (
    <div style={{ textAlign: 'center', maxWidth: 1300 }}>
      {words.map((w, i) => {
        const start = delay + i * perWord
        const o = interpolate(frame, [start, start + 10], [0, 1], {
          extrapolateLeft: 'clamp',
          extrapolateRight: 'clamp',
        })
        const y = interpolate(frame, [start, start + 10], [24, 0], {
          extrapolateLeft: 'clamp',
          extrapolateRight: 'clamp',
        })
        return (
          <span
            key={i}
            style={{
              display: 'inline-block',
              fontFamily: font,
              fontWeight: weight,
              fontSize: size,
              letterSpacing: tracking,
              color,
              opacity: o,
              transform: `translateY(${y}px)`,
              marginRight: size * 0.26,
              textShadow: glow ? '0 0 40px rgba(99,102,241,0.55)' : undefined,
            }}
          >
            {w}
          </span>
        )
      })}
    </div>
  )
}

/** Small uppercase eyebrow line. */
export const Eyebrow: React.FC<{ text: string; delay?: number; font: string }> = ({ text, delay = 0, font }) => {
  const frame = useCurrentFrame()
  const o = interpolate(frame, [delay, delay + 12], [0, 1], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })
  return (
    <div
      style={{
        fontFamily: font,
        fontWeight: 600,
        fontSize: 26,
        letterSpacing: 8,
        textTransform: 'uppercase',
        color: 'rgba(129,140,248,0.9)',
        opacity: o,
        marginBottom: 28,
        textAlign: 'center',
      }}
    >
      {text}
    </div>
  )
}
