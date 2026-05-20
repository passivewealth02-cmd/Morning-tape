import { ImageResponse } from 'next/og'

export const alt = 'Maintena — AI Property Maintenance Software'
export const size = { width: 1200, height: 630 }
export const contentType = 'image/png'

export default function OgImage() {
  return new ImageResponse(
    (
      <div
        style={{
          width: '100%',
          height: '100%',
          display: 'flex',
          flexDirection: 'column',
          justifyContent: 'space-between',
          padding: '72px',
          background: 'linear-gradient(135deg, #1e1b4b 0%, #4f46e5 60%, #6366f1 100%)',
          color: 'white',
          fontFamily: 'sans-serif',
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: 20 }}>
          <div
            style={{
              width: 72,
              height: 72,
              borderRadius: 18,
              background: 'white',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
            }}
          >
            <svg width="42" height="42" viewBox="0 0 120 120" fill="none">
              <path
                d="M 24 96 L 24 30 L 60 70 L 96 30 L 96 96"
                stroke="#4f46e5"
                strokeWidth={16}
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
          </div>
          <div style={{ fontSize: 40, fontWeight: 700 }}>Maintena</div>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 24 }}>
          <div style={{ fontSize: 64, fontWeight: 800, lineHeight: 1.1, maxWidth: 900 }}>
            AI Property Maintenance Software
          </div>
          <div style={{ fontSize: 32, color: 'rgba(255,255,255,0.85)', maxWidth: 900, lineHeight: 1.3 }}>
            Capture requests, triage with AI, dispatch vendors, and track every repair to completion.
          </div>
        </div>

        <div style={{ fontSize: 26, color: 'rgba(255,255,255,0.7)' }}>trymaintena.com</div>
      </div>
    ),
    size,
  )
}
