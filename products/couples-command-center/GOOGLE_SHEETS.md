# Relationship & Couples Command Center™ — Google Sheets Setup

How to turn the `.xlsx` master into the shareable **"Make a Copy"** Google Sheets
template that buyers receive.

---

## 1. Upload & convert

1. Go to [drive.google.com](https://drive.google.com) → **New → File upload** →
   choose `Couples_Command_Center.xlsx`.
2. Once uploaded, double-click it → **Open with → Google Sheets**.
3. **File → Save as Google Sheets**. A native Sheets copy is created.

---

## 2. Check it converted cleanly

Google Sheets reads everything this workbook uses (`SUM`, `COUNTA`, `COUNTIF`, `MIN`,
`MAX`, `ABS`, `AVERAGE`, `IF`, `IFERROR`, named ranges, data validation, conditional
formatting, charts). Confirm:

- [ ] **Shared Bills** totals **$3,908**.
- [ ] **Fair Share** shows **40% / 60%** income shares, a 50/50 split leaving **$2,246 /
      $4,346** with a **$2,100 gap**, and a proportional split leaving both at **62.8%**.
      Those two percentages **must be identical** — that's the whole point of the page,
      and it's the block buyers screenshot.
- [ ] **Invisible Labour** shows **23.0** vs **8.0** hours, a **2.88×** split and **780
      extra hours a year**.
- [ ] **Money Goals** shows 5 goals all on track or better; **Savings** totals **$940** a
      month; **Date Nights** shows 4; **Weekly Check-In** shows 4; **Big Conversations**
      shows 6 had and 2 still to have.
- [ ] **Dashboard** shows all 12 KPIs, the How-You're-Doing bars, the hours donut, the
      saved-by-month chart and **Together Score 90%**.
- [ ] Your names propagate — several header cells pull `=PartnerA` and `=PartnerB` from
      **Settings**, so changing them there changes them everywhere.
- [ ] Dropdowns work (who pays, bill category, goal status, how often, yes/no).

> If a named range looks off, open **Data → Named ranges** and confirm the names
> (`SharedBills`, `ShareA`, `ShareB`, `LeftHalfA`, `LeftHalfB`, `HalfGap`, `FairA`,
> `FairB`, `KeepPctA`, `KeepPctB`, `Fairness`, `HoursA`, `HoursB`, `ChoreRatio`,
> `ExtraHoursYear`, `HealthRange`, etc.) resolve — Sheets occasionally re-scopes on
> import.

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

> **Tell buyers to share the Sheet with each other.** In Google Sheets this is genuinely
> better than the Excel version, because both partners can open the same live file. That's
> a real selling point — mention it.

---

## 5. Three things to say plainly in the listing

> **This is an organizing tool, not relationship advice or counselling.** It cannot tell
> anyone what a fair life together looks like — only the two of them can. Say that. It's
> honest, it keeps you clear of Etsy's Services policy, and it stops anyone expecting
> therapy from a spreadsheet.

> **Both partners fill in the labour page separately, then compare.** This instruction is
> in the file and on the printable, and it matters: almost every couple finds the columns
> don't match, and the point is the conversation, not the scoreboard. Put it in your
> description — it's also the line that shows buyers you've thought about this properly.

> **There's a safeguarding line in the file, and it should stay.** If money is being
> controlled rather than shared, a spreadsheet is not the right help. You don't need to
> put that in the listing, but don't remove it from the workbook.

> Keep the delivery a plain digital download. It's fine to answer a buyer's question in
> Messages after a sale, but never advertise support, setup, **couples coaching or
> counselling**, or "free updates" as part of the listing.
