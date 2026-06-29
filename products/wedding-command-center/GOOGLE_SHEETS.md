# Wedding Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the 32-sheet Excel
flagship. Same tab order: **Dashboard, Couple, Timeline, Checklist,
Budget, Payments, Savings, Vendors, Comparison, Contracts, Guests, RSVP,
Seating, BridalParty, Attire, Beauty, Ceremony, Reception, Menu, Music,
Floral, Decor, ShotList, Honeymoon, Registry, Gifts, ThankYou,
VisionBoard, Packing, WeddingDay, Analytics, Settings**.

> Build **Settings** first — it defines every named range the rest of
> the system depends on.

---

## 1. Settings (named ranges)

| Cell | Named range | Default |
| ---- | ----------- | ------- |
| `C6` | `WeddingDate` | `2027-06-12` |
| `C7` | `TotalBudget` | `45000` |
| `C8` | `EstGuests` | `140` |
| `C9` | `NumTables` | `16` |
| `C10` | `SeatsPerTable` | `10` |
| `C11` | `MonthlySavings` | `1500` |
| `C12` | `TaxRate` | `0.08` |
| `C13` | `TodayDate` | `=TODAY()` |

Dropdown list named ranges (Data → Named ranges):

`BudgetCatList, VendorTypeList, StatusList, MealList, PriorityList,
StyleList, RsvpList, ContractList, PayMethodList, RelationList,
YesNoList, PhaseList` — see Settings cols E–J, rows 7+ and 27+.

---

## 2. Cross-sheet named ranges (define these once)

| Range | Points to |
| ----- | --------- |
| `ChkTask` | `Checklist!B5:B124` |
| `ChkDone` | `Checklist!F5:F124` |
| `ChkDue` | `Checklist!E5:E124` |
| `TLTask` | `Timeline!B5:B84` |
| `TLDone` | `Timeline!F5:F84` |
| `BudCat` | `Budget!A5:A25` |
| `BudPlanned` | `Budget!B5:B25` |
| `BudActual` | `Budget!C5:C25` |
| `BudPaid` | `Budget!E5:E25` |
| `BudTotalActual` | `Budget!C26` |
| `BudTotalPaid` | `Budget!E26` |
| `PayBalance` | `Payments!D5:D44` |
| `PayPaid` | `Payments!F5:F44` |
| `PayVendor` | `Payments!A5:A44` |
| `VenName` | `Vendors!B5:B34` |
| `VenContract` | `Vendors!I5:I34` |
| `GuestName` | `Guests!A5:A204` |
| `GuestRsvp` | `Guests!G5:G204` |
| `GuestSeats` | `Guests!H5:H204` |
| `GuestInvited` | `Guests!F5:F204` |
| `GuestMeal` | `Guests!I5:I204` |

---

## 3. Key automation formulas

### Timeline — auto-anchored target dates

Each milestone is offset from the wedding date. Type the offset (months)
in a helper column, then:

```sheets
D5 =ARRAYFORMULA(IF(B5:B84="","",WeddingDate-Helper5:Helper84*30))
E5 =ARRAYFORMULA(IF(B5:B84="","",
      IF(F5:F84="Yes","Complete",
        IF(D5:D84<TODAY(),"Overdue","In Progress"))))
```

### Checklist — status + completion

```sheets
G5 =ARRAYFORMULA(IF(B5:B124="","",
      IF(F5:F124="Yes","Complete",
        IF(E5:E124<TODAY(),"Overdue","In Progress"))))
```

Completion % (used on Dashboard):

```sheets
=IFERROR(COUNTIF(ChkDone,"Yes")/COUNTA(ChkTask),0)
```

### Budget — paid, remaining, totals

```sheets
E5 =ARRAYFORMULA(IF(A5:A25="","",D5:D25))        ' Paid = Deposit
F5 =ARRAYFORMULA(IF(A5:A25="","",C5:C25-E5:E25)) ' Remaining
C26 =SUM(C5:C25)   ' total estimated
E26 =SUM(E5:E25)   ' total paid
```

### Payments — late alert

```sheets
I5 =ARRAYFORMULA(IF(E5:E44="","",E5:E44-TODAY()))   ' Days Out
```

Conditional format (whole row): `=AND($E5<>"",$E5<TODAY(),$F5<>"Yes")`
→ red. Within 14 days & unpaid → gold.

### RSVP — live counts

```sheets
Invited seats  =SUMPRODUCT((GuestInvited="Yes")*IFERROR(GuestSeats,0))
Accepted seats =SUMPRODUCT((GuestRsvp="Accepted")*IFERROR(GuestSeats,0))
Declined       =COUNTIF(GuestRsvp,"Declined")
Pending        =COUNTIF(GuestRsvp,"Pending")
RSVP %         =IFERROR((COUNTIF(GuestRsvp,"Accepted")
                +COUNTIF(GuestRsvp,"Declined"))/COUNTA(GuestName),0)
Attendance %   =IFERROR(
                 SUMPRODUCT((GuestRsvp="Accepted")*IFERROR(GuestSeats,0))
                 /SUMPRODUCT((GuestInvited="Yes")*IFERROR(GuestSeats,0)),0)
```

Meal counts (per meal type):

```sheets
=SUMPRODUCT((GuestMeal="Beef")*(GuestRsvp="Accepted")*IFERROR(GuestSeats,0))
```

Or with QUERY for the whole table at once:

```sheets
=QUERY(Guests!G5:I204,
  "select I, count(I) where G='Accepted' and I is not null group by I
   label count(I) 'Guests'",0)
```

### Seating — interactive table assignment

**Step 1 — assign guests.** On the **Guests** tab, column **P "Table"** has a
dropdown (`TableList` = `Settings!B17:B36`, i.e. Table 1…Table 20). Pick a
table for each guest. Add the named range `GuestTable → Guests!P5:P204`.

**Step 2 — rosters auto-fill.** On the **Seating** tab, each table row pulls
its guest list and seat count automatically:

```sheets
A5  Table 1   (typed Table 1…Table 20 down column A)
B5  =TEXTJOIN(", ", TRUE, IF(GuestTable=$A5, GuestName, ""))   ← live roster
C5  =SUMPRODUCT((GuestTable=$A5)*IFERROR(GuestSeats,0))        ← seats used
D5  10                                                          ← capacity
E5  =IF(D5="","",D5-C5)                                         ← open seats
```

(In Google Sheets the `TEXTJOIN(...IF(range...))` evaluates as an array
automatically — no Ctrl+Shift+Enter needed.)

Conditional formatting: over-capacity `=AND($C5<>"",$C5>$D5)` → red,
full `=$E5=0` → mint, VIP `=ISNUMBER(SEARCH("VIP",$F5))` → gold.

**Step 3 — Table Inspector (“see who’s sitting here”).** A dropdown cell
`InspectTable` (`Seating!I5`, validated to `TableList`) drives a live panel:

```sheets
I6  =SUMPRODUCT((GuestTable=InspectTable)*IFERROR(GuestSeats,0))      ← seated
I7  =IFERROR(INDEX($D$5:$D$24,MATCH(InspectTable,$A$5:$A$24,0)),"")   ← capacity
I8  =IF(I7="","",I7-I6)                                               ← open
H11 =TEXTJOIN(", ",TRUE,IF(GuestTable=InspectTable,
        GuestName&" ("&IFERROR(GuestSeats,0)&")",""))                 ← roster
I19 =SUMPRODUCT((GuestTable=InspectTable)*(GuestMeal="Beef")
        *IFERROR(GuestSeats,0))   (repeat per meal type)              ← meal counts
```

Capacity summary also shows **Guests Not Yet Seated**:
`=SUMPRODUCT((GuestRsvp="Accepted")*(GuestTable="")*IFERROR(GuestSeats,0))`.

> Note: true hover-tooltips need a macro (breaks in Sheets/mobile). This
> Inspector + the always-on roster column are the no-macro, cross-platform
> equivalent — pick a table and instantly see exactly who's seated there.

### Analytics — readiness score

```sheets
Budget Health       =IFERROR(1-MAX(BudTotalActual-TotalBudget,0)/TotalBudget,0)
Planning Completion =IFERROR(COUNTIF(ChkDone,"Yes")/COUNTA(ChkTask),0)
Timeline Health     =IFERROR(COUNTIF(TLDone,"Yes")/COUNTA(TLTask),0)
Vendor Completion   =IFERROR(COUNTIF(VenContract,"Signed")/COUNTA(VenName),0)
Guest Completion    =IFERROR((COUNTIF(GuestRsvp,"Accepted")
                     +COUNTIF(GuestRsvp,"Declined"))/COUNTA(GuestName),0)
Payment Health      =IFERROR(COUNTIF(PayPaid,"Yes")/COUNTA(PayVendor),0)

Overall Readiness   =AVERAGE(of the six scores above)
```

---

## 4. Dashboard KPI block

```sheets
Days Until Wedding   =MAX(WeddingDate-TODAY(),0)
Budget Remaining     =TotalBudget-BudTotalPaid
Budget Spent         =BudTotalActual
Planning Complete    =IFERROR(COUNTIF(ChkDone,"Yes")/COUNTA(ChkTask),0)
RSVP Complete        =IFERROR((COUNTIF(GuestRsvp,"Accepted")
                      +COUNTIF(GuestRsvp,"Declined"))/COUNTA(GuestName),0)
Vendors Booked       =COUNTIF(VenContract,"Signed")
Payments Outstanding =SUMPRODUCT((PayPaid<>"Yes")*IFERROR(PayBalance,0))
Guest Count          =SUMPRODUCT((GuestRsvp="Accepted")*IFERROR(GuestSeats,0))
Tables Filled        =ROUNDUP(<guest count>/SeatsPerTable,0)&" / "&NumTables
Tasks Due This Week  =SUMPRODUCT((ChkDue<>"")*(ChkDue-TODAY()>=0)
                      *(ChkDue-TODAY()<=7)*(ChkDone<>"Yes"))
```

### Quick-navigation chips

Insert a row of cells, then for each: right-click → **Insert link** →
**Sheets in this spreadsheet** → pick the target tab. Style: fill
`#1B4F48`, white bold, centered. This reproduces the Excel hyperlinked
nav.

---

## 5. Vision Board — image-in-cell

For each inspiration tile:

- **Insert → Image → Image in cell** (the cell becomes a drag-drop
  image holder).
- Use a 16-tile grid (Venue, Dress, Florals, Cake, Hair, Makeup, Decor,
  Tablescapes, Stationery, Color Palette, Bouquets, Photo Poses,
  Reception Setup, Lighting, Favors, Travel Ideas).
- Below each tile keep Notes / Estimated cost / Priority / Vendor /
  Purchase-status rows.

(Excel equivalent: **Insert → Pictures → Place in Cell**.)

---

## 6. Apps Script — daily refresh + weekly snapshot

```javascript
function refreshDaily() { SpreadsheetApp.flush(); }

function weeklySnapshot() {
  const ss = SpreadsheetApp.getActive();
  const log = ss.getSheetByName('Reports') || ss.insertSheet('Reports');
  const v = (s, c) => ss.getSheetByName(s).getRange(c).getValue();
  log.appendRow([
    new Date(),
    v('Dashboard','B6'),   // days until
    v('Dashboard','J6'),   // RSVP %
    v('Analytics','F6'),   // readiness
  ]);
}

function installTriggers() {
  ScriptApp.newTrigger('refreshDaily').timeBased().everyDays(1).atHour(5).create();
  ScriptApp.newTrigger('weeklySnapshot').timeBased().everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
}
```

---

## 7. Brand palette (luxury edition)

| Token | Hex | Use |
| ----- | --- | --- |
| Primary | `#1B4F48` | Headers, KPI values, nav chips, time column |
| Accent (Gold) | `#937356` | KPI labels, section sub-heads, dividers |
| Gold Light | `#C9A86A` | Divider rows under each page header |
| Surface | `#E5D3BA` | Input cells, totals row |
| Highlight | `#75E6C1` | Complete / on-track / positive states |
| Blush | `#F3E4DD` | Vision-board image tiles |
| Ivory | `#FBF8F2` | Readiness score panel |
| Danger | `#C94C4C` | Overdue / over-capacity / declined |
| Muted Row | `#F4ECDE` | Alternating row stripes |

Every page uses the two-row header (title + italic subtitle) with a
**gold divider row** beneath — that's the signature luxury treatment.
Header rows fill `#1B4F48` white bold all-caps; alternating data rows
white / `#F4ECDE`; input cells fill `#E5D3BA`.
