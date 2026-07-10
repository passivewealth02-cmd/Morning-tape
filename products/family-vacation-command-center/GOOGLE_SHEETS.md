# Family Vacation Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Family Profile, Master Trip, Budget, Savings,
Itinerary, Transport, Hotels, Packing, Kids, Meals, Activities, Reservations,
Documents, Emergency, Road Trip, Responsibilities, Points, Photo Vault, Journal,
Post-Trip, Wishlist, Analytics, Settings**.

> Build **Settings** first (trip details + dropdown lists), then the Budget,
> Packing, Reservations & Documents, then the Dashboard + Analytics. Add the
> named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `FamilyName`, `TripName`, `Destination`, `TripType`, `TripStart`
(date), `TripEnd` (date), `TravelerCount` (5), `SavingsGoal` (8500).

Lists: `TripTypeList, ExpCatList, MemberList, ActTypeList, PackCatList,
TransTypeList, DocTypeList, PriorityList, YesNoList, PayStatusList,
DocStatusList, ResTypeList, RoleList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `TripBudget` | `Budget!B15` (planned total) | `PackStatus` | `Packing!E5:E64` |
| `SpentTotal` | `Budget!C15` (actual total) | `ResConfirm` | `Reservations!F5:F34` |
| `SavedTotal` | `Budget!H10` (fund total) | `DocStatus` | `Documents!F5:F20` |
| `ExpCat` | `Budget!A5:A14` | `ActName` | `Activities!A5:A44` |
| `ExpPlanned` | `Budget!B5:B14` | `ReadinessRange` | `Analytics!C7:C10` |
| `ExpActual` | `Budget!C5:C14` | `SaveVal` | `Analytics!C16:C21` |

---

## 3. Budget & Savings

```sheets
Remaining (per row)  =B5-C5
Total budget         =SUM(B5:B14)
Money saved          =SUM(H5:H9)
Budget remaining     =TripBudget-SpentTotal
Cost / person        =IFERROR(TripBudget/TravelerCount,0)
Cost / day           =IFERROR(TripBudget/(TripEnd-TripStart+1),0)
Savings progress     =IFERROR(MIN(SavedTotal/SavingsGoal,1),0)
```

Data-bar the savings progress; flag negative "Remaining" cells red.

---

## 4. Dashboard — the 12 KPIs

```sheets
Trip Countdown     =MAX(TripStart-TODAY(),0)
Total Budget       =TripBudget
Money Saved        =SavedTotal
Budget Remaining   =TripBudget-SpentTotal
Travelers          =TravelerCount
Days Traveling     =TripEnd-TripStart+1
Packing Complete   =IFERROR(COUNTIF(PackStatus,"Yes")/COUNTA(PackStatus),0)
Bookings Done      =IFERROR(COUNTIF(ResConfirm,"Yes")/COUNTA(ResConfirm),0)
Documents Ready    =IFERROR(COUNTIF(DocStatus,"Ready")/COUNTA(DocStatus),0)
Activities Planned =COUNTA(ActName)
Cost / Person      =IFERROR(TripBudget/TravelerCount,0)
Readiness Score    =IFERROR(AVERAGE(ReadinessRange),0)
```

Charts: Budget by Category (donut), Vacation Fund Growth (line), Trip Readiness
(bar), Planned vs Actual (column). Turn off auto data labels.

---

## 5. Analytics — Family Trip Readiness Score

```sheets
Savings progress    =IFERROR(MIN(SavedTotal/SavingsGoal,1),0)
Bookings confirmed  =IFERROR(COUNTIF(ResConfirm,"Yes")/COUNTA(ResConfirm),0)
Packing complete    =IFERROR(COUNTIF(PackStatus,"Yes")/COUNTA(PackStatus),0)
Documents ready     =IFERROR(COUNTIF(DocStatus,"Ready")/COUNTA(DocStatus),0)
Readiness Score     =IFERROR(AVERAGE(C7:C10),0)
```

Power features: `ARRAYFORMULA`, `QUERY` ("spend by category", "unpacked items"),
`FILTER`/`SORT`/`UNIQUE`, all wrapped in `IFERROR`.

Photo & Memory Vault: click a cell ▸ Insert ▸ Image ▸ **Image in cell**, or
paste `=IMAGE("your-photo-link")`.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, status color flags. Keep it
premium and consistent — that polish is what makes it feel like a travel app,
not a spreadsheet.
