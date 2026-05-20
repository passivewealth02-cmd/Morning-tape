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
              width: 64,
              height: 64,
              borderRadius: 16,
              background: 'white',
              color: '#4f46e5',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              fontSize: 44,
              fontWeight: 800,
            }}
          >
            M
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
