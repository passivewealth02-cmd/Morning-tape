# Dating Life Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Dating_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `COUNTA`, `COUNTIF`, `MIN`,
`MAX`, `AVERAGE`, `IF`, `IFERROR`, named ranges, data validation, conditional formatting,
charts). Confirm:

- [ ] **Dating Funnel** shows 240 → 68 → 14 → 5 → 2, with **5.8%** match-to-date and
      **35.7%** second-date rate.
- [ ] It also shows **69 hours**, **$496.98**, **$35.50** per first date and **$99.40 and
      13.8 hours** per second date. This is the block buyers screenshot; make sure it
      survives the import.
- [ ] **Effort & Reciprocity** totals **48** vs **16**, gives **3.0×**, and the verdict
      cell reads **CARRYING ALL OF IT**. Gaps of 5 or more shade red.
- [ ] **Green & Red Flags** shows **9 green**, **2 red**, **net 7**, with green "Yes"
      mint and red "Yes" red.
- [ ] **Conversations** shades the four rows under 35% — all four are people never met.
- [ ] **Safety Plan** shows **9** steps held.
- [ ] **Dashboard** shows all 12 KPIs, the How-It's-Going bars, the hours donut, the
      first-dates chart and **Dating Score 90%**.
- [ ] Dropdowns work (status, met on, how it felt, weight, yes/no) — from **Settings**.

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`Matches`, `SecondDateRate`, `HoursSpent`, `SpendTotal`, `CostPerSecondDate`,
> `HoursPerSecondDate`, `YourEffort`, `TheirEffort`, `EffortRatio`, `GreenCount`,
> `RedCount`, `FlagNet`, `SafetyHeld`, `HealthRange`, etc.) resolve — Sheets occasionally
> re-scopes on import.

---

## 3. Make the "Make a Copy" share link

1. **Share → General access → Anyone with the link → Viewer.**
2. Copy the link. It ends in `/edit?usp=sharing`.
3. **Replace `/edit?usp=sharing` with `/copy`.**

```
https://docs.google.com/spreadsheets/d/FILE_ID/copy
```

Now anyone who clicks it gets **"Make a copy"** — their own private copy, your master
untouched. Put this link in `GOOGLE_SHEETS_TEMPLATE_LINK.txt` in the buyer's download.

---

## 4. What buyers receive

- The `.xlsx` file (opens in Excel & Google Sheets)
- The **"Make a Copy"** Google Sheets link (1 click → their own copy)
- The 12-page printable PDF
- A Start-Here quick-start guide

---

## 5. Three things to say plainly in the listing

> **This is an organizing and reflection tool, not advice.** Not relationship advice, not
> psychological advice, not a method for finding love. Say that in the listing. It is
> honest, it keeps you well clear of Etsy's Services policy, and — genuinely — it is the
> framing that makes buyers trust the product.

> **The scores are a mirror, not a verdict.** A low effort ratio doesn't mean someone is a
> bad person; it means the buyer is doing more of the work, and only she can decide what
> to do about that. The workbook says this out loud on the Effort tab and the printable
> repeats it. Repeat it in your description too — it's also, not coincidentally, the most
> shareable line in the whole listing.

> **Tell buyers it's private.** People will hesitate to buy something that logs their
> dating life. Say plainly: it's a file on your own Drive or computer, nobody else sees
> it, and it isn't an app or an account. The in-file label already reads "private ·
> yours".

> Keep the delivery a plain digital download. It's fine to answer a buyer's question in
> Messages after a sale, but never advertise support, setup, **dating coaching, profile
> reviews**, or "free updates" as part of the listing — Etsy's Services policy treats that
> as selling a service, and this niche is watched closely because it's saturated with
> coaching.
