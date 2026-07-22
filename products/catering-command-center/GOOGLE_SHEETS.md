# Catering Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Plate Costing, Menu Packages, Event Quotes, Staffing,
Rentals, Bookings, Inventory, Waste Log, Ordering, Cash & Deposits, Clients,
Settings**.

> Build **Plate Costing** and **Menu Packages** first (the engine), then Event
> Quotes and the rest, then the Dashboard. Add the named ranges below
> (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Company`, `Owner`, `TargetFC` (0.30), `MarginGoal` (0.30), `LaborLimit`
(0.40), `BookingGoal` (6), `Currency`.

Lists: `UnitList, PackageList, StatusList, YesNoList`.

---

## 2. Plate Costing & Menu Packages — the engine

```sheets
Plate     Cost per head =SUM(component costs)  → PlatedCostHead
Package   Margin/head   =Price/head − Cost/head
          Food %        =Cost/head ÷ Price/head
```

Named: `PlatedCostHead, PkgItem, PkgCostHead, PkgPriceHead, PkgMargin`. The Plated
Dinner package pulls its cost with `=PlatedCostHead`.

---

## 3. Event Quotes — quote = full event P&L

```sheets
Price/Hd   =IFERROR(INDEX(PkgPriceHead,MATCH(Package,PkgItem,0)),0)
Cost/Hd    =IFERROR(INDEX(PkgCostHead,MATCH(Package,PkgItem,0)),0)
Revenue    =Guests*Price/Hd + Service
Food Cost  =Guests*Cost/Hd
Margin $   =Revenue − Food − Staff − Rentals
Margin %   =IFERROR(Margin$/Revenue,0)
```

Named: `EventName, EventPkg, EventGuests, EventStaff, EventRev, EventFood,
EventMarginPct`, plus totals `TotalRevenue, TotalFood`, and summary cells
`FoodCostPct` (=TotalFood/TotalRevenue), `AvgEventMargin` (=AVERAGE(EventMarginPct)),
`LaborPct` (=SUM(EventStaff)/TotalRevenue).

---

## 4. Dashboard — the 12 KPIs

```sheets
Events         =COUNTA(EventName)
Avg Guests     =AVERAGE(EventGuests)
Revenue        =TotalRevenue
Avg Per Head   =TotalRevenue/SUM(EventGuests)
Food Cost      =FoodCostPct
Top Package    =INDEX(EventPkg,MATCH(MAX(EventRev),EventRev,0))
Avg Event      =TotalRevenue/COUNTA(EventName)
Avg Margin     =AvgEventMargin
Labor          =LaborPct
Packages       =COUNTA(PkgItem)
Waste %        =IFERROR(WasteTotal/TotalRevenue,0)
Catering Score =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Revenue by Event (column) from the Event Quotes tab.

---

## 5. Catering Score (6 dimensions)

```sheets
Food cost on target    =IFERROR(MIN(TargetFC/FoodCostPct,1),0)
Margin per event       =IFERROR(MIN(AvgEventMargin/MarginGoal,1),0)
Packages fully costed   =IFERROR(COUNTIF(PkgCostHead,">0")/COUNTA(PkgItem),0)
Labor under control     =IFERROR(1-MIN(LaborPct/LaborLimit,1),0)
Bookings vs goal        =IFERROR(MIN(COUNTA(EventName)/BookingGoal,1),0)
Gross margin            =IFERROR(MIN((1-FoodCostPct)/0.7,1),0)
Catering Score          =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `SUMPRODUCT`, `INDEX`, `MATCH`, `COUNTIF`, `COUNTA`, `AVERAGE`,
`MAX`, `MIN`, `IFERROR`, color scales (food cost %, catering score) and
conditional formatting (booking status, low inventory).

---

## 6. Printables

The 12-page PDF is print-ready as-is (US Letter) — a plate cost card, package
price list, event quote & run sheet, staffing sheet & a client contact sheet for
every event. Print any tab: File ▸ Print ▸ fit to width.

> A business tool, not financial or accounting advice — confirm figures with your
> own books.

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
