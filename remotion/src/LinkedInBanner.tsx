import React from 'react'
import { AbsoluteFill } from 'remotion'
import { BRAND, FONT } from './brand'

// LinkedIn personal banner: 1584 x 396.
// Keep the lower-left ~360px clearer — the profile photo overlaps there.
export const LinkedInBanner: React.FC = () => {
  return (
    <AbsoluteFill style={{ background: '#07090F' }}>
      <AbsoluteFill
        style={{
          background:
            'radial-gradient(circle at 78% 40%, rgba(79,70,229,0.28) 0%, rgba(79,70,229,0.08) 40%, rgba(0,0,0,0) 65%)',
        }}
      />
      {/* faint grid */}
      <AbsoluteFill
        style={{
          backgroundImage:
            'linear-gradient(rgba(255,255,255,0.04) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.04) 1px, transparent 1px)',
          backgroundSize: '48px 48px',
          maskImage: 'linear-gradient(90deg, transparent 0%, black 55%)',
        }}
      />

      <div
        style={{
          position: 'absolute',
          left: 470,
          top: 96,
          display: 'flex',
          flexDirection: 'column',
          gap: 18,
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: 20 }}>
          <div
            style={{
              width: 84,
              height: 84,
              borderRadius: 20,
              background: BRAND.indigo,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              boxShadow: '0 0 50px rgba(79,70,229,0.5)',
            }}
          >
            <svg viewBox="0 0 120 120" width={46} height={46} fill="none">
              <path
                d="M 24 96 L 24 30 L 60 70 L 96 30 L 96 96"
                stroke="#fff"
                strokeWidth={16}
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
          </div>
          <span style={{ fontFamily: FONT, fontWeight: 600, fontSize: 68, letterSpacing: -2, color: '#fff' }}>
            Maintena
          </span>
        </div>

        <div style={{ fontFamily: FONT, fontWeight: 600, fontSize: 40, letterSpacing: -0.5, color: 'rgba(255,255,255,0.92)' }}>
          Stop losing maintenance requests and chasing vendors.
        </div>
        <div style={{ fontFamily: FONT, fontWeight: 500, fontSize: 27, color: 'rgba(165,180,252,0.95)' }}>
          The AI operations layer for property maintenance · trymaintena.com
        </div>
      </div>
    </AbsoluteFill>
  )
}
