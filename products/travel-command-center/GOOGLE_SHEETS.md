# Travel Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Trip Profile, Itinerary, Budget, Expenses,
Flights, Hotels, Transport, Activities, Restaurants, Packing, Documents,
Currency, Savings, Road Trip, Group Travel, Checklists, Memories, Journal,
Analytics, Settings**.

> Build **Settings** first (trip details + dropdown lists), then the trackers,
> then the Dashboard. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `TripName` (C6), `DepartDate` (C7), `ReturnDate` (C8),
`Travelers` (C9), `HomeCurrency` (C10), `TripCurrency` (C11),
`FXRate` (C12), `TotalBudget` (C13), `SavingsGoal` (C14), `TripDays` (C15).

Lists: `CatList, StatusList, YesNoList, ExpenseCatList, FlightStatusList,
DocTypeList, TransportList, ActivityTypeList, CurrencyList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `DepartDate` | `Settings!C7` | `FlightAirline` | `Flights!A5:A24` |
| `TripDays` | `Settings!C15` | `HotelName` | `Hotels!A5:A20` |
| `TripCountries` | `'Trip Profile'!B15:B20` | `ActStatus` | `Activities!H5:H30` |
| `ItinDate` | `Itinerary!B5:B…` | `DocName` | `Documents!A5:A24` |
| `ItinCity` | `Itinerary!C5:C…` | `DocReady` | `Documents!F5:F24` |
| `ExpDate` | `Expenses!A5:A…` | `PackDone` | `Packing!<hidden col>` |
| `ExpCat` | `Expenses!B5:B…` | `BudgetCat` | `Budget!A5:A20` |
| `ExpAmount` | `Expenses!E5:E…` | `BudgetActual` | `Budget!C5:C20` |
| `BudgetTotalPlanned` | `Budget!B21` | `BudgetTotalActual` | `Budget!C21` |
| `SavProgress` | `Savings!C9` | | |

---

## 3. Budget — actuals pull from Expenses automatically

```sheets
Actual (per category)  =SUMIF(ExpCat, <category cell>, ExpAmount)
Remaining              =Planned-Actual
% Used                 =IFERROR(Actual/Planned,0)
Budget / Day           =IFERROR(BudgetTotalPlanned/TripDays,0)
```

Add a data-bar conditional format on the **% Used** column (green → gold as it
approaches 100%).

---

## 4. Dashboard — the 12 KPIs

```sheets
Days to Departure  =MAX(DepartDate-TODAY(),0)
Total Budget       =BudgetTotalPlanned
Budget Left        =BudgetTotalPlanned-BudgetTotalActual
Spent So Far       =BudgetTotalActual
Flights Booked     =COUNTA(FlightAirline)
Hotels Confirmed   =COUNTA(HotelName)
Activities Booked  =COUNTIF(ActStatus,"Booked")
Countries          =COUNTA(TripCountries)
Docs Ready         =IFERROR(COUNTIF(DocReady,"Yes")/MAX(COUNTA(DocName),1),0)
Packing            =IFERROR(COUNTIF(PackDone,"Yes")/MAX(COUNTA(PackDone),1),0)
Savings Goal       =SavProgress
Readiness          =IFERROR(AVERAGE(Analytics!C7:C11),0)
```

Charts: Spending by Category (donut), Budget vs Actual (column), Readiness
(bar), Savings Progress (donut). Turn off auto data labels.

---

## 5. Itinerary — auto day numbering & countdowns

```sheets
Day #        =IF(B5="","",B5-DepartDate+1)
Days Away    =IF(B5="","",B5-TODAY())
```

Power features: `ARRAYFORMULA`, `QUERY` ("next 3 activities"),
`FILTER`/`SORT`, all wrapped in `IFERROR`. Currency conversion on the
**Currency** tab: `=Amount*FXRate`.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, data-bar budget rows.
Keep it warm, premium and consistent — that polish is the whole feel.
