import React from 'react'
import { interpolate, spring, useCurrentFrame, useVideoConfig } from 'remotion'
import { BRAND, FONT } from '../brand'
import { AmbientGlow, CINE, Eyebrow, RevealText, Scene } from './helpers'

/* ------------------------------------------------------------------ */
/* 1. COLD OPEN — the problem, whispered in the dark                   */
/* ------------------------------------------------------------------ */

const problems = [
  { label: 'Missed call — Unit 4B', x: -560, y: -220, at: 6 },
  { label: 'Unread text', x: 520, y: -180, at: 12 },
  { label: 'Lost email thread', x: -480, y: 230, at: 18 },
  { label: '"Any update?"', x: 460, y: 250, at: 24 },
  { label: 'Vendor no-show', x: 0, y: -290, at: 30 },
]

export const CineColdOpen: React.FC<{ duration: number }> = ({ duration }) => {
  const frame = useCurrentFrame()
  return (
    <Scene duration={duration} from={1.0} to={1.07}>
      <AmbientGlow x={50} y={40} size={55} />
      {problems.map((p, i) => {
        const o = interpolate(frame, [p.at, p.at + 12, p.at + 50, p.at + 66], [0, 0.5, 0.5, 0.12], {
          extrapolateLeft: 'clamp',
          extrapolateRight: 'clamp',
        })
        const drift = Math.sin((frame + i * 25) / 34) * 7
        return (
          <div
            key={p.label}
            style={{
              position: 'absolute',
              transform: `translate(${p.x}px, ${p.y + drift}px)`,
              opacity: o,
              fontFamily: FONT,
              fontWeight: 500,
              fontSize: 30,
              color: CINE.textDim,
              border: `1px solid ${CINE.glassBorder}`,
              background: CINE.glass,
              padding: '12px 26px',
              borderRadius: 999,
            }}
          >
            {p.label}
          </div>
        )
      })}
      <RevealText
        font={FONT}
        text="Every day, requests slip through the cracks."
        delay={26}
        perWord={5}
        size={72}
        color="#fff"
      />
    </Scene>
  )
}

/* ------------------------------------------------------------------ */
/* 2. THE QUESTION — the turn                                          */
/* ------------------------------------------------------------------ */

export const CineQuestion: React.FC<{ duration: number }> = ({ duration }) => {
  return (
    <Scene duration={duration} from={1.02} to={1.08}>
      <AmbientGlow x={50} y={55} size={70} />
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8 }}>
        <RevealText font={FONT} text="What if maintenance" delay={4} perWord={5} size={84} color="rgba(255,255,255,0.85)" />
        <RevealText font={FONT} text="ran itself?" delay={22} perWord={6} size={110} weight={700} color="#A5B4FC" glow />
      </div>
    </Scene>
  )
}

/* ------------------------------------------------------------------ */
/* 3. THE SCAN — one gesture                                           */
/* ------------------------------------------------------------------ */

const QR_PATTERN = [
  '111111101011011111111',
  '100000100110001000001',
  '101110101101001011101',
  '101110100011001011101',
  '101110101110001011101',
  '100000100101001000001',
  '111111101010101111111',
  '000000000000000000000',
  '110101110110110101101',
  '001011001011011100010',
  '110010110101110011011',
]

export const CineScan: React.FC<{ duration: number }> = ({ duration }) => {
  const frame = useCurrentFrame()
  const { fps } = useVideoConfig()
  const qrIn = spring({ frame: frame - 4, fps, config: { damping: 18 }, durationInFrames: 22 })
  const scan = interpolate(frame % 48, [0, 48], [4, 96])
  const lock = spring({ frame: frame - 34, fps, config: { damping: 13 }, durationInFrames: 18 })
  const cell = 13

  return (
    <Scene duration={duration}>
      <AmbientGlow x={42} y={50} size={55} />
      <div style={{ display: 'flex', alignItems: 'center', gap: 90 }}>
        <div style={{ position: 'relative', transform: `scale(${qrIn})` }}>
          <div
            style={{
              background: '#fff',
              padding: 22,
              borderRadius: 20,
              boxShadow: '0 0 90px rgba(79,70,229,0.35)',
            }}
          >
            {QR_PATTERN.map((row, y) => (
              <div key={y} style={{ display: 'flex' }}>
                {row.split('').map((v, x) => (
                  <div key={x} style={{ width: cell, height: cell, background: v === '1' ? '#0B0E1A' : 'transparent' }} />
                ))}
              </div>
            ))}
            <div
              style={{
                position: 'absolute',
                left: 10,
                right: 10,
                height: 5,
                top: `${scan}%`,
                background: BRAND.indigoLight,
                boxShadow: `0 0 26px ${BRAND.indigoLight}`,
                borderRadius: 4,
              }}
            />
          </div>
        </div>

        <div style={{ opacity: lock, transform: `translateX(${(1 - lock) * -26}px)` }}>
          <div
            style={{
              border: `1px solid rgba(129,140,248,0.35)`,
              background: 'rgba(79,70,229,0.12)',
              borderRadius: 22,
              padding: '26px 34px',
              boxShadow: '0 0 60px rgba(79,70,229,0.18)',
            }}
          >
            <div style={{ fontFamily: FONT, fontWeight: 600, fontSize: 24, color: '#A5B4FC', letterSpacing: 2 }}>
              LOCATION CONFIRMED
            </div>
            <div style={{ fontFamily: FONT, fontWeight: 700, fontSize: 40, color: '#fff', marginTop: 8 }}>
              Maple Gardens · Unit 4B
            </div>
          </div>
          <div style={{ marginTop: 18, fontFamily: FONT, fontSize: 26, color: CINE.textDim }}>
            No app. No login. One scan.
          </div>
        </div>
      </div>
      <div style={{ position: 'absolute', bottom: 150 }}>
        <RevealText font={FONT} text="A tenant scans their door." delay={8} size={54} color="rgba(255,255,255,0.9)" />
      </div>
    </Scene>
  )
}

/* ------------------------------------------------------------------ */
/* 4. THE INTELLIGENCE — AI triage                                     */
/* ------------------------------------------------------------------ */

const aiLabels = [
  { text: 'PLUMBING', color: '#A5B4FC', border: 'rgba(129,140,248,0.5)', at: 22 },
  { text: 'EMERGENCY', color: '#FCA5A5', border: 'rgba(248,113,113,0.5)', at: 32 },
  { text: 'DISPATCH: PLUMBER', color: 'rgba(255,255,255,0.85)', border: CINE.glassBorder, at: 42 },
]

export const CineTriage: React.FC<{ duration: number }> = ({ duration }) => {
  const frame = useCurrentFrame()
  const { fps } = useVideoConfig()
  const cardIn = spring({ frame: frame - 2, fps, config: { damping: 17 }, durationInFrames: 22 })
  const title = 'Water leaking under kitchen sink'
  const typed = title.slice(0, Math.round(interpolate(frame, [6, 30], [0, title.length], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })))

  return (
    <Scene duration={duration}>
      <AmbientGlow x={58} y={45} size={60} />
      <div
        style={{
          transform: `scale(${cardIn})`,
          width: 920,
          background: CINE.glass,
          border: `1px solid ${CINE.glassBorder}`,
          borderRadius: 28,
          padding: 48,
          backdropFilter: 'blur(8px)',
        }}
      >
        <div style={{ fontFamily: FONT, fontWeight: 600, fontSize: 22, color: CINE.textFaint, letterSpacing: 3 }}>
          INCOMING REQUEST
        </div>
        <div style={{ fontFamily: FONT, fontWeight: 700, fontSize: 44, color: '#fff', marginTop: 14, minHeight: 60 }}>
          {typed}
          <span style={{ opacity: frame % 16 < 8 ? 1 : 0, color: '#A5B4FC' }}>|</span>
        </div>
        <div style={{ display: 'flex', gap: 16, marginTop: 34 }}>
          {aiLabels.map(l => {
            const s = spring({ frame: frame - l.at, fps, config: { damping: 11 }, durationInFrames: 16 })
            return (
              <span
                key={l.text}
                style={{
                  transform: `scale(${s})`,
                  fontFamily: FONT,
                  fontWeight: 700,
                  fontSize: 26,
                  letterSpacing: 3,
                  padding: '14px 28px',
                  borderRadius: 999,
                  color: l.color,
                  border: `1px solid ${l.border}`,
                  background: 'rgba(255,255,255,0.04)',
                }}
              >
                {l.text}
              </span>
            )
          })}
        </div>
      </div>
      <div style={{ position: 'absolute', bottom: 150 }}>
        <RevealText font={FONT} text="AI reads it. Grades it. Routes it. In seconds." delay={26} perWord={4} size={54} />
      </div>
    </Scene>
  )
}

/* ------------------------------------------------------------------ */
/* 5. THE DISPATCH                                                     */
/* ------------------------------------------------------------------ */

export const CineDispatch: React.FC<{ duration: number }> = ({ duration }) => {
  const frame = useCurrentFrame()
  const { fps } = useVideoConfig()
  const slide = spring({ frame: frame - 2, fps, config: { damping: 19 }, durationInFrames: 26 })
  const check = spring({ frame: frame - 24, fps, config: { damping: 9 }, durationInFrames: 16 })

  return (
    <Scene duration={duration}>
      <AmbientGlow x={50} y={50} size={55} hue="22,163,74" />
      <div
        style={{
          transform: `translateY(${(1 - slide) * 60}px)`,
          opacity: slide,
          width: 880,
          background: CINE.glass,
          border: `1px solid ${CINE.glassBorder}`,
          borderRadius: 28,
          padding: 44,
          display: 'flex',
          alignItems: 'center',
          gap: 30,
        }}
      >
        <div
          style={{
            width: 96,
            height: 96,
            borderRadius: 24,
            background: 'rgba(79,70,229,0.25)',
            border: '1px solid rgba(129,140,248,0.4)',
            color: '#A5B4FC',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            fontFamily: FONT,
            fontWeight: 700,
            fontSize: 38,
          }}
        >
          MP
        </div>
        <div style={{ flex: 1 }}>
          <div style={{ fontFamily: FONT, fontWeight: 700, fontSize: 40, color: '#fff' }}>Mike&apos;s Plumbing Co.</div>
          <div style={{ fontFamily: FONT, fontSize: 27, color: CINE.textDim, marginTop: 6 }}>
            Top-rated · Available now · 12 min away
          </div>
        </div>
        <span
          style={{
            transform: `scale(${check})`,
            fontFamily: FONT,
            fontWeight: 700,
            fontSize: 30,
            padding: '14px 30px',
            borderRadius: 999,
            color: '#86EFAC',
            border: '1px solid rgba(74,222,128,0.45)',
            background: 'rgba(22,163,74,0.15)',
            boxShadow: '0 0 44px rgba(22,163,74,0.25)',
          }}
        >
          DISPATCHED ✓
        </span>
      </div>
      <div style={{ position: 'absolute', bottom: 150 }}>
        <RevealText font={FONT} text="The right pro. Already on the way." delay={20} size={54} />
      </div>
    </Scene>
  )
}

/* ------------------------------------------------------------------ */
/* 6. THE CALM — everyone in the loop                                  */
/* ------------------------------------------------------------------ */

const steps = ['Received', 'Assigned', 'In progress', 'Completed']

export const CineLoop: React.FC<{ duration: number }> = ({ duration }) => {
  const frame = useCurrentFrame()
  const progress = interpolate(frame, [8, 52], [0, 3], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })

  return (
    <Scene duration={duration}>
      <AmbientGlow x={50} y={42} size={55} />
      <div style={{ display: 'flex', alignItems: 'flex-start' }}>
        {steps.map((s, i) => {
          const done = progress >= i
          const active = Math.floor(progress) === i && !done
          return (
            <React.Fragment key={s}>
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', width: 230 }}>
                <div
                  style={{
                    width: 72,
                    height: 72,
                    borderRadius: 999,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    fontFamily: FONT,
                    fontWeight: 700,
                    fontSize: 30,
                    color: done || active ? '#fff' : CINE.textFaint,
                    background: done ? 'rgba(22,163,74,0.8)' : active ? BRAND.indigo : 'rgba(255,255,255,0.06)',
                    border: `1px solid ${done || active ? 'transparent' : CINE.glassBorder}`,
                    boxShadow: done ? '0 0 34px rgba(22,163,74,0.35)' : active ? '0 0 34px rgba(79,70,229,0.4)' : 'none',
                  }}
                >
                  {done ? '✓' : i + 1}
                </div>
                <div
                  style={{
                    marginTop: 18,
                    fontFamily: FONT,
                    fontWeight: 600,
                    fontSize: 27,
                    color: done || active ? '#fff' : CINE.textFaint,
                  }}
                >
                  {s}
                </div>
              </div>
              {i < steps.length - 1 && (
                <div style={{ width: 110, height: 4, background: 'rgba(255,255,255,0.1)', marginTop: 34, position: 'relative', borderRadius: 4 }}>
                  <div
                    style={{
                      position: 'absolute',
                      inset: 0,
                      borderRadius: 4,
                      width: `${interpolate(progress - i, [0, 1], [0, 100], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })}%`,
                      background: 'linear-gradient(90deg, #4F46E5, #22C55E)',
                    }}
                  />
                </div>
              )}
            </React.Fragment>
          )
        })}
      </div>
      <div style={{ position: 'absolute', bottom: 150 }}>
        <RevealText font={FONT} text="And no one ever has to ask for an update again." delay={22} perWord={4} size={54} />
      </div>
    </Scene>
  )
}

/* ------------------------------------------------------------------ */
/* 7. FINALE — logo, tagline, call to action                           */
/* ------------------------------------------------------------------ */

export const CineFinale: React.FC<{ duration: number }> = ({ duration }) => {
  const frame = useCurrentFrame()
  const { fps } = useVideoConfig()
  const pop = spring({ frame: frame - 4, fps, config: { damping: 14 }, durationInFrames: 26 })
  const draw = interpolate(frame, [8, 34], [1, 0], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })
  const word = interpolate(frame, [28, 44], [0, 1], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })
  const tag = interpolate(frame, [46, 60], [0, 1], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })
  const url = interpolate(frame, [58, 72], [0, 1], { extrapolateLeft: 'clamp', extrapolateRight: 'clamp' })

  return (
    <Scene duration={duration} from={1.0} to={1.04}>
      <AmbientGlow x={50} y={48} size={65} />
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 30, transform: `scale(${pop})` }}>
          <div
            style={{
              width: 170,
              height: 170,
              borderRadius: 40,
              background: BRAND.indigo,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              boxShadow: '0 0 110px rgba(79,70,229,0.55)',
            }}
          >
            <svg viewBox="0 0 120 120" width={86} height={86} fill="none">
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
          <span
            style={{
              fontFamily: FONT,
              fontWeight: 600,
              fontSize: 108,
              letterSpacing: -4,
              color: '#fff',
              opacity: word,
              transform: `translateX(${(1 - word) * 22}px)`,
            }}
          >
            Maintena
          </span>
        </div>
        <div
          style={{
            marginTop: 44,
            fontFamily: FONT,
            fontWeight: 500,
            fontSize: 40,
            letterSpacing: -0.5,
            color: 'rgba(255,255,255,0.75)',
            opacity: tag,
          }}
        >
          Property maintenance, on autopilot.
        </div>
        <div
          style={{
            marginTop: 24,
            fontFamily: FONT,
            fontWeight: 600,
            fontSize: 34,
            color: '#A5B4FC',
            opacity: url,
            textShadow: '0 0 30px rgba(99,102,241,0.5)',
          }}
        >
          trymaintena.com — start free
        </div>
      </div>
    </Scene>
  )
}
