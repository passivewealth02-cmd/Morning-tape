const { ImageResponse } = require('next/og')
const React = require('react')
const fs = require('fs')
const h = React.createElement

const INDIGO_BG = 'linear-gradient(135deg, #1e1b4b 0%, #4f46e5 60%, #6366f1 100%)'

// ---------- Avatar 400x400 (shown as a circle) ----------
function Avatar() {
  return h(
    'div',
    {
      style: {
        width: '100%', height: '100%', display: 'flex',
        alignItems: 'center', justifyContent: 'center',
        background: INDIGO_BG, fontFamily: 'sans-serif',
      },
    },
    h(
      'div',
      {
        style: {
          width: 220, height: 220, borderRadius: 48, background: 'white',
          color: '#4f46e5', display: 'flex', alignItems: 'center',
          justifyContent: 'center', fontSize: 150, fontWeight: 800,
        },
      },
      'M'
    )
  )
}

// ---------- Banner 1500x500 ----------
function Banner() {
  return h(
    'div',
    {
      style: {
        width: '100%', height: '100%', display: 'flex', flexDirection: 'column',
        justifyContent: 'center', alignItems: 'center', textAlign: 'center',
        background: INDIGO_BG, color: 'white', fontFamily: 'sans-serif', padding: 60,
      },
    },
    h(
      'div',
      { style: { display: 'flex', alignItems: 'center', gap: 24, marginBottom: 28 } },
      h(
        'div',
        {
          style: {
            width: 84, height: 84, borderRadius: 22, background: 'white', color: '#4f46e5',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            fontSize: 60, fontWeight: 800,
          },
        },
        'M'
      ),
      h('div', { style: { fontSize: 64, fontWeight: 800 } }, 'Maintena')
    ),
    h(
      'div',
      { style: { fontSize: 40, fontWeight: 700, maxWidth: 1100 } },
      'AI Property Maintenance Software'
    ),
    h(
      'div',
      { style: { fontSize: 26, color: 'rgba(255,255,255,0.82)', marginTop: 14, maxWidth: 1000 } },
      'Capture requests · AI triage · Dispatch vendors · Track every repair'
    ),
    h(
      'div',
      { style: { fontSize: 24, color: 'rgba(255,255,255,0.7)', marginTop: 26 } },
      'trymaintena.com'
    )
  )
}

async function render(node, opts, outfile) {
  const res = new ImageResponse(node, opts)
  const buf = Buffer.from(await res.arrayBuffer())
  fs.writeFileSync(outfile, buf)
  console.log('wrote', outfile, buf.length, 'bytes')
}

;(async () => {
  await render(h(Avatar), { width: 400, height: 400 }, 'public/brand/twitter-avatar.png')
  await render(h(Banner), { width: 1500, height: 500 }, 'public/brand/twitter-banner.png')
})()
