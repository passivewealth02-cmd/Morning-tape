"use client";

import React, { useState, useEffect } from "react";
import { AlertCircle } from "lucide-react";

export default function HomePage() {
  const [tickers, setTickers] = useState("AAPL, NVDA, BTC, TSLA");
  const [brief, setBrief] = useState<any>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [time, setTime] = useState(new Date());

  useEffect(() => {
    const t = setInterval(() => setTime(new Date()), 1000);
    return () => clearInterval(t);
  }, []);

  const today = new Date().toLocaleDateString("en-US", {
    weekday: "long",
    month: "long",
    day: "numeric",
    year: "numeric",
  });

  const generateBrief = async () => {
    if (!tickers.trim()) return;
    setLoading(true);
    setError(null);
    setBrief(null);

    try {
      const res = await fetch("/api/generate-brief", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ tickers, today }),
      });
      if (!res.ok) throw new Error("Request failed");
      const data = await res.json();
      setBrief(data);
    } catch (err) {
      console.error(err);
      setError("Couldn't generate brief. Try again in a moment.");
    } finally {
      setLoading(false);
    }
  };

  const Sparkline = ({ data, color }: { data: number[]; color: string }) => {
    if (!data || data.length === 0) return null;
    const w = 60, h = 20;
    const min = Math.min(...data), max = Math.max(...data);
    const range = max - min || 1;
    const points = data
      .map((v, i) => `${(i / (data.length - 1)) * w},${h - ((v - min) / range) * h}`)
      .join(" ");
    return (
      <svg width={w} height={h} className="inline-block">
        <polyline points={points} fill="none" stroke={color} strokeWidth="1.2" strokeLinejoin="round" />
      </svg>
    );
  };

  const handleSubscribe = async (tier: "trader" | "professional") => {
    try {
      const res = await fetch("/api/stripe/checkout", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ tier }),
      });
      const data = await res.json();
      if (data.url) window.location.href = data.url;
    } catch (err) {
      console.error(err);
      alert("Couldn't start checkout. Try again.");
    }
  };

  return (
    <div className="min-h-screen grain-paper">
      <header className="max-w-6xl mx-auto px-8 pt-8 pb-4">
        <div className="flex items-end justify-between flex-wrap gap-2 pb-2 mono text-[10px] uppercase tracking-[0.2em] text-ink-dim">
          <div>Vol. III · No. 247</div>
          <div className="hidden sm:block">{today}</div>
          <div>One Dollar at the Newsstand</div>
        </div>
        <div className="rule mb-2" />
        <div className="rule-thin mb-6" />
        <div className="text-center py-4">
          <div className="mono text-[11px] uppercase tracking-[0.4em] text-ink-dim mb-3">
            ⁂ Established for the Discerning Trader ⁂
          </div>
          <h1 className="display text-7xl md:text-9xl leading-none" style={{ fontWeight: 500, letterSpacing: "-0.04em" }}>
            The Morning Tape
          </h1>
          <div className="display italic text-2xl text-ink-dim mt-3" style={{ fontWeight: 300 }}>
            A Pre-Market Brief, Delivered Before the Bell
          </div>
        </div>
        <div className="rule-thin mt-6 mb-2" />
        <div className="rule" />
        <div className="flex justify-between items-center pt-3 mono text-[10px] uppercase tracking-[0.2em] text-ink-dim">
          <div>Equities · Crypto · Macro</div>
          <div className="flex items-center gap-2"><span className="blink">●</span> Live Edition</div>
          <div className="hidden sm:block">{time.toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit" })} ET</div>
        </div>
      </header>

      <section className="max-w-6xl mx-auto px-8 py-16 drop">
        <div className="grid md:grid-cols-12 gap-8 items-start">
          <div className="md:col-span-8">
            <div className="mono text-[10px] uppercase tracking-[0.25em] text-accent mb-4 small-caps">⸻ The Premise ⸻</div>
            <h2 className="display text-6xl md:text-7xl leading-[0.95] mb-8" style={{ fontWeight: 400, letterSpacing: "-0.025em" }}>
              Forty browser tabs,<br/>
              <span className="italic" style={{ fontWeight: 300 }}>distilled to one cup of coffee.</span>
            </h2>
            <p className="body-text text-xl leading-[1.55] drop-cap" style={{ fontWeight: 400 }}>
              Every morning at seven, while you're still pouring the first cup, we deliver a personalized briefing. Overnight moves on every ticker you watch. Macro releases that will move tape. The three things that actually matter — written sharp enough to read in five minutes, deep enough that you walk into the open with conviction.
            </p>
          </div>
          <aside className="md:col-span-4 bg-paper-dark p-6 border border-ink">
            <div className="mono text-[10px] uppercase tracking-[0.2em] text-ink-dim mb-3">Today's Subscriber Count</div>
            <div className="display text-5xl mb-4" style={{ fontWeight: 600 }}>3,847</div>
            <div className="rule-thin mb-4" />
            <div className="space-y-2 mono text-[11px] text-ink-dim">
              <div className="flex justify-between"><span>Avg. Read Time</span><span className="text-ink">4m 12s</span></div>
              <div className="flex justify-between"><span>Open Rate</span><span className="text-ink">68.3%</span></div>
              <div className="flex justify-between"><span>NPS</span><span className="text-ink">71</span></div>
            </div>
            <div className="rule-thin my-4" />
            <p className="body-text italic text-sm leading-snug">"I cancelled three other newsletters after my first week."</p>
            <p className="mono text-[10px] text-ink-dim mt-2">— M.K., Hedge Fund PM</p>
          </aside>
        </div>
      </section>

      <section className="max-w-6xl mx-auto px-8 py-12">
        <div className="rule-double mb-6" />
        <div className="text-center mb-12">
          <div className="mono text-[10px] uppercase tracking-[0.3em] text-accent mb-3">⸻ Section II ⸻</div>
          <h3 className="display text-5xl md:text-6xl" style={{ fontWeight: 400, letterSpacing: "-0.02em" }}>Generate Your Edition</h3>
          <p className="body-text italic text-lg text-ink-dim mt-3">Enter your watchlist below. We'll write your briefing live.</p>
        </div>
        <div className="bg-paper-dark p-8 mb-8 border border-ink">
          <label className="mono text-[10px] uppercase tracking-[0.2em] text-ink-dim block mb-3">⁘ Watchlist ⁘</label>
          <div className="flex flex-col sm:flex-row gap-3">
            <input
              type="text"
              value={tickers}
              onChange={(e) => setTickers(e.target.value)}
              placeholder="AAPL, NVDA, BTC, TSLA"
              className="flex-1 mono text-sm px-4 py-3 focus:outline-none placeholder:opacity-40 bg-paper border border-ink"
            />
            <button
              onClick={generateBrief}
              disabled={loading}
              className="px-8 py-3 mono text-xs uppercase tracking-[0.2em] disabled:opacity-50 transition-all hover:opacity-80 bg-ink text-paper border border-ink"
            >
              {loading ? "Setting Type..." : "Print Edition →"}
            </button>
          </div>
          <p className="body-text italic text-sm text-ink-dim mt-4">Stocks, crypto, ETFs, indices — separated by commas.</p>
        </div>
        {error && (
          <div className="p-5 mb-8 flex items-start gap-3 border border-accent" style={{ background: "#F4E5DC" }}>
            <AlertCircle size={18} className="text-accent mt-0.5 flex-shrink-0" />
            <span className="body-text text-accent">{error}</span>
          </div>
        )}
        {brief && (
          <article className="drop bg-paper-light border border-ink p-10">
            <div className="text-center mb-8">
              <div className="rule-double mb-3" />
              <div className="mono text-[10px] uppercase tracking-[0.3em] text-ink-dim mb-2">Morning Edition · {today}</div>
              <div className="rule-thin" />
            </div>
            <div className="text-center mb-10">
              <div className="mono text-[10px] uppercase tracking-[0.3em] text-accent mb-3 small-caps">⸻ {brief.kicker} ⸻</div>
              <h1 className="display text-5xl md:text-6xl leading-[0.95] mx-auto max-w-3xl" style={{ fontWeight: 500, letterSpacing: "-0.02em" }}>{brief.headline}</h1>
              <div className="mono text-[10px] uppercase tracking-[0.2em] text-ink-dim mt-6">By The Editors · Filed {time.toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit" })} ET</div>
            </div>
            <div className="rule-thin mb-8" />
            <div className="max-w-3xl mx-auto mb-12">
              <p className="body-text text-xl leading-[1.6] drop-cap" style={{ fontWeight: 400 }}>{brief.lede}</p>
            </div>
            <div className="rule-thin mb-8" />
            <div className="mb-12">
              <div className="text-center mb-6">
                <div className="mono text-[10px] uppercase tracking-[0.3em] text-ink-dim mb-2">⁘ The Issues ⁘</div>
                <h3 className="display text-3xl italic" style={{ fontWeight: 400 }}>Your Watchlist</h3>
              </div>
              <div className="space-y-0">
                {brief.watchlist?.map((item: any, i: number) => (
                  <div key={i} className="grid md:grid-cols-12 gap-4 py-5" style={{ borderTop: i === 0 ? "1px solid #1A1612" : "0.5px solid rgba(26,22,18,0.3)", borderBottom: i === brief.watchlist.length - 1 ? "1px solid #1A1612" : "none" }}>
                    <div className="md:col-span-3 flex items-center gap-3">
                      <span className="mono text-2xl" style={{ fontWeight: 500 }}>{item.ticker}</span>
                      <Sparkline data={item.spark} color={item.direction === "up" ? "#2D5016" : "#8B0000"} />
                    </div>
                    <div className="md:col-span-2 mono text-sm">
                      <div className={item.direction === "up" ? "text-green-mark" : "text-accent"} style={{ fontWeight: 500 }}>{item.move}</div>
                      <div className="text-ink-dim text-xs mt-1">{item.level}</div>
                    </div>
                    <div className="md:col-span-7">
                      <p className="body-text text-base leading-snug mb-2">{item.thesis}</p>
                      <div className="body-text italic text-sm text-ink-dim">
                        <span className="mono text-[10px] uppercase tracking-widest text-accent not-italic">Trigger · </span>
                        {item.trigger}
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
            <div className="mb-12 max-w-2xl mx-auto">
              <div className="text-center mb-6">
                <div className="mono text-[10px] uppercase tracking-[0.3em] text-ink-dim mb-2">⁘ The Diary ⁘</div>
                <h3 className="display text-3xl italic" style={{ fontWeight: 400 }}>Today's Calendar</h3>
              </div>
              <div className="rule mb-3" />
              <div className="space-y-0">
                {brief.calendar?.map((event: any, i: number) => (
                  <div key={i} className="grid grid-cols-12 gap-3 py-3 items-baseline" style={{ borderBottom: "0.5px solid rgba(26,22,18,0.2)" }}>
                    <div className="col-span-3 mono text-xs text-ink-dim">{event.time}</div>
                    <div className="col-span-7 body-text" style={{ fontWeight: 500 }}>
                      {event.event}
                      {event.consensus && <span className="body-text italic text-ink-dim text-sm ml-2">— consensus: {event.consensus}</span>}
                    </div>
                    <div className={`col-span-2 text-right mono text-[10px] uppercase tracking-widest ${event.weight === "high" ? "text-accent" : "text-ink-dim"}`}>
                      {event.weight === "high" ? "★★★" : event.weight === "medium" ? "★★" : "★"}
                    </div>
                  </div>
                ))}
              </div>
            </div>
            <div className="rule-double mb-8" />
            <div className="max-w-2xl mx-auto text-center">
              <div className="mono text-[10px] uppercase tracking-[0.3em] text-accent mb-4 small-caps">⸻ The Editor's Take ⸻</div>
              <p className="display italic text-2xl leading-[1.4]" style={{ fontWeight: 400 }}>"{brief.thesis}"</p>
            </div>
            <div className="rule-double mt-8" />
            <div className="text-center mt-8 mono text-[10px] uppercase tracking-[0.3em] text-ink-dim">⁂ End of Edition · See You at the Bell ⁂</div>
          </article>
        )}
      </section>

      <section className="max-w-6xl mx-auto px-8 py-16">
        <div className="rule-double mb-6" />
        <div className="text-center mb-12">
          <div className="mono text-[10px] uppercase tracking-[0.3em] text-accent mb-3">⸻ Section III ⸻</div>
          <h3 className="display text-5xl md:text-6xl" style={{ fontWeight: 400, letterSpacing: "-0.02em" }}>Subscription Terms</h3>
          <p className="body-text italic text-lg text-ink-dim mt-3">Two tiers. Cancel anytime. No fine print.</p>
        </div>
        <div className="grid md:grid-cols-2 gap-0 border border-ink">
          <div className="p-10 border-r border-ink">
            <div className="mono text-[10px] uppercase tracking-[0.3em] text-ink-dim mb-2">⸻ Standard ⸻</div>
            <h4 className="display text-3xl mb-1" style={{ fontWeight: 500 }}>The Trader</h4>
            <p className="body-text italic text-ink-dim mb-6">For the daily speculator</p>
            <div className="rule-thin mb-6" />
            <div className="flex items-baseline gap-2 mb-8">
              <span className="display text-7xl" style={{ fontWeight: 400 }}>$29</span>
              <span className="mono text-xs text-ink-dim uppercase tracking-widest">per mensem</span>
            </div>
            <ul className="space-y-3 mb-8 body-text">
              {["Daily morning edition, 7am ET", "Up to 25 watchlist tickers", "Web archive of past editions", "Macro calendar integration", "Equities, crypto & ETFs"].map((f, i) => (
                <li key={i} className="flex items-baseline gap-3">
                  <span className="mono text-xs text-ink-dim flex-shrink-0">{String(i+1).padStart(2,'0')}</span>
                  <span>{f}</span>
                </li>
              ))}
            </ul>
            <button onClick={() => handleSubscribe("trader")} className="w-full py-4 mono text-xs uppercase tracking-[0.25em] transition-all hover:opacity-80 border border-ink">Subscribe — $29 →</button>
          </div>
          <div className="p-10 relative bg-ink text-paper">
            <div className="absolute top-3 right-3 mono text-[10px] uppercase tracking-[0.3em]">★ Recommended</div>
            <div className="mono text-[10px] uppercase tracking-[0.3em] mb-2 opacity-60">⸻ Premium ⸻</div>
            <h4 className="display text-3xl mb-1" style={{ fontWeight: 500 }}>The Professional</h4>
            <p className="body-text italic mb-6 opacity-60">For serious capital</p>
            <div className="border-t border-paper/30 mb-6" />
            <div className="flex items-baseline gap-2 mb-8">
              <span className="display text-7xl" style={{ fontWeight: 400 }}>$49</span>
              <span className="mono text-xs uppercase tracking-widest opacity-60">per mensem</span>
            </div>
            <ul className="space-y-3 mb-8 body-text">
              {["Everything in Standard", "Unlimited watchlist", "Mid-day update at noon ET", "After-hours summary", "Earnings preview the night before", "SMS alerts on major moves"].map((f, i) => (
                <li key={i} className="flex items-baseline gap-3">
                  <span className="mono text-xs flex-shrink-0 opacity-50">{String(i+1).padStart(2,'0')}</span>
                  <span>{f}</span>
                </li>
              ))}
            </ul>
            <button onClick={() => handleSubscribe("professional")} className="w-full py-4 mono text-xs uppercase tracking-[0.25em] transition-all hover:opacity-90 bg-paper text-ink">Subscribe — $49 →</button>
          </div>
        </div>
      </section>

      <footer className="max-w-6xl mx-auto px-8 pb-12">
        <div className="rule-double mb-6" />
        <div className="text-center mono text-[10px] uppercase tracking-[0.25em] text-ink-dim space-y-2">
          <div>The Morning Tape · A Daily Pre-Market Brief</div>
          <div>Set in Fraunces & Crimson Pro · Published Before the Bell</div>
          <div className="body-text italic text-ink-dim text-sm pt-3">Not financial advice. Markets are risk. Trade your conviction, not ours.</div>
        </div>
        <div className="rule mt-6" />
      </footer>
    </div>
  );
}
