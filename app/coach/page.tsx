import type { Metadata } from 'next'
import { CoachApp } from '@/components/coach/coach-app'

export const metadata: Metadata = {
  title: 'Design Autopilot Coach — Etsy PNG Design Recipes',
  description:
    'A beginner-friendly design coach for Etsy PNG sellers. Get a layout, font pairing, colour palette, object placement map, design warnings and a copy-paste AI prompt from one simple form.',
  robots: { index: false, follow: false },
}

export default function CoachPage() {
  return <CoachApp />
}
