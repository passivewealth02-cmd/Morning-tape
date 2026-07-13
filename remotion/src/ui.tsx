import React from 'react'
import { interpolate, useCurrentFrame } from 'remotion'
import { BRAND, FONT } from './brand'

/** A caption that fades + slides up. */
export const Caption: React.FC<{ children: React.ReactNode; delay?: number; color?: string }> = ({
  children,
  delay = 6,
  color = BRAND.ink,
}) => {
  const frame = useCurrentFrame()
  const o = interpolate(frame, [delay, delay + 12], [0, 1], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })
  const y = interpolate(frame, [delay, delay + 12], [16, 0], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })
  return (
    <div
      style={{
        fontFamily: FONT,
        fontWeight: 600,
        fontSize: 52,
        letterSpacing: -1.5,
        color,
        textAlign: 'center',
        opacity: o,
        transform: `translateY(${y}px)`,
        maxWidth: 1200,
      }}
    >
      {children}
    </div>
  )
}

/** A small pill/chip. */
export const Chip: React.FC<{
  label: string
  bg?: string
  color?: string
  style?: React.CSSProperties
}> = ({ label, bg = BRAND.gray100, color = BRAND.gray600, style }) => (
  <span
    style={{
      fontFamily: FONT,
      fontWeight: 600,
      fontSize: 30,
      padding: '10px 22px',
      borderRadius: 999,
      background: bg,
      color,
      whiteSpace: 'nowrap',
      ...style,
    }}
  >
    {label}
  </span>
)
