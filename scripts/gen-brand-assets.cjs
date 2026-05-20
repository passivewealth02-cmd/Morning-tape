const { ImageResponse } = require('next/og')
const React = require('react')
const fs = require('fs')
const h = React.createElement

const DEEP_BG = 'radial-gradient(130% 130% at 50% 22%, #3b35b0 0%, #1e1b4b 52%, #0a0a18 100%)'
const GLOW = 'radial-gradient(circle, rgba(99,102,241,0.55) 0%, rgba(99,102,241,0) 70%)'

// Custom rounded geometric "M" monogram
function Monogram({ size, id }) {
  return h(
    'svg',
    { width: size, height: size, viewBox: '0 0 120 120', fill: 'none' },
    h(
      'defs',
      null,
      h(
        'linearGradient',
        { id, x1: '0', y1: '0', x2: '1', y2: '1' },
        h('stop', { offset: '0%', stopColor: '#ffffff' }),
        h('stop', { offset: '100%', stopColor: '#a5b4fc' })
      )
    ),
    h('path', {
      d: 'M 24 96 L 24 30 L 60 70 L 96 30 L 96 96',
      stroke: `url(#${id})`,
      strokeWidth: 16,
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
      fill: 'none',
    })
  )
}

// ---------- Avatar 400x400 (shown as a circle) ----------
function Avatar() {
  return h(
    'div',
    {
      style: {
        width: '100%', height: '100%', position: 'relative', display: 'flex',
        alignItems: 'center', justifyContent: 'center', background: DEEP_BG, fontFamily: 'sans-serif',
      },
    },
    h('div', { style: { position: 'absolute', width: 300, height: 300, borderRadius: 9999, background: GLOW, top: 50, left: 50, display: 'flex' } }),
    h('div', { style: { display: 'flex', position: 'relative' } }, h(Monogram, { size: 244, id: 'av' }))
  )
}

// ---------- Banner 1500x500 ----------
function Banner() {
  return h(
    'div',
    {
      style: {
        width: '100%', height: '100%', position: 'relative', display: 'flex',
        flexDirection: 'column', justifyContent: 'center', alignItems: 'center',
        background: DEEP_BG, color: 'white', fontFamily: 'sans-serif',
      },
    },
    h('div', { style: { position: 'absolute', width: 900, height: 900, borderRadius: 9999, background: GLOW, top: -250, left: 300, display: 'flex' } }),
    h(
      'div',
      { style: { display: 'flex', alignItems: 'center', gap: 26, position: 'relative' } },
      h(Monogram, { size: 132, id: 'bn' }),
      h('div', { style: { width: 2, height: 78, background: 'rgba(255,255,255,0.18)' } }),
      h('div', { style: { fontSize: 78, fontWeight: 700, letterSpacing: -2 } }, 'Maintena')
    ),
    h(
      'div',
      {
        style: {
          fontSize: 27, color: 'rgba(226,232,255,0.78)', marginTop: 30,
          letterSpacing: 5, textTransform: 'uppercase', position: 'relative',
        },
      },
      'AI Property Maintenance Software'
    )
  )
}

async function render(node, opts, outfile) {
  const res = new ImageResponse(node, opts)
  fs.writeFileSync(outfile, Buffer.from(await res.arrayBuffer()))
  console.log('wrote', outfile)
}

;(async () => {
  await render(h(Avatar), { width: 400, height: 400 }, 'public/brand/twitter-avatar.png')
  await render(h(Banner), { width: 1500, height: 500 }, 'public/brand/twitter-banner.png')
})()
