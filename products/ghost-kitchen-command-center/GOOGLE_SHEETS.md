# Ghost Kitchen Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Item Margin, Menu & Margins, Platform P&L, Virtual Brands,
Packaging, Order Volume, Inventory, Waste Log, Ordering, Payouts, Promotions,
Settings**.

> Build **Item Margin** and **Menu & Margins** first (the engine), then Platform
> P&L and the rest, then the Dashboard. Add the named ranges below
> (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Kitchen`, `Owner`, `TargetFC` (0.30), `MarginGoal` (0.40), `CommLimit`
(0.40), `DirectGoal` (0.15), `Commission` (0.25), `Currency`.

Lists: `UnitList, PlatformList, StatusList, YesNoList`.

---

## 2. Item Margin & Menu & Margins — the engine

```sheets
Item Margin   App menu price  (input)
              App commission  =-Price*Commission
              Food cost       (input, negative)
              Packaging       (input, negative)
              NET MARGIN      =SUM(above)              → SigItemNet
Menu          Net Margin      =AppPrice*(1-Commission)-Food-Packaging
              Net %           =NetMargin/AppPrice
              Food %          =Food/AppPrice
```

Named: `SigItemNet, MenuItem, MenuPrice, MenuFood, MenuNet, MenuNetPct,
MenuFoodPct`, plus `AvgNetPct` (=AVERAGE(MenuNetPct)) and `FoodCostPct`
(=AVERAGE(MenuFoodPct)).

---

## 3. Platform P&L

```sheets
Gross        =Orders*AvgOrder
Net Payout   =Gross*(1-Commission)
TOTAL gross  =SUM(PlatGross) ; TOTAL net =SUM(PlatNet)
BlendedComm  =(TotalGross-NetPayout)/TotalGross
DirectShare  =Direct orders/TotalOrders
AvgOrder     =TotalGross/TotalOrders
```

Named: `PlatName, PlatOrders, PlatGross, PlatNet, TotalOrders, TotalGross,
BlendedComm, NetPayout, DirectShare, AvgOrder`. Virtual Brands adds `BrandName,
BrandOrders, BrandRev`; Waste adds `WasteTotal`.

---

## 4. Dashboard — the 12 KPIs

```sheets
Menu Items       =COUNTA(MenuItem)
Avg App Price    =AVERAGE(MenuPrice)
Food Cost        =FoodCostPct
Avg Net Margin   =AvgNetPct
Blended Comm     =BlendedComm
Top Item         =INDEX(MenuItem,MATCH(MAX(MenuNet),MenuNet,0))
Weekly Orders    =TotalOrders
Weekly Revenue   =TotalGross
Net Payout       =NetPayout
Avg Order        =AvgOrder
Virtual Brands   =COUNTA(BrandName)
Kitchen Score    =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Net Payout by Platform (column) from the Platform P&L tab.

---

## 5. Kitchen Score (6 dimensions)

```sheets
Food cost on target    =IFERROR(MIN(TargetFC/FoodCostPct,1),0)
Net margin healthy     =IFERROR(MIN(AvgNetPct/MarginGoal,1),0)
Menu fully costed       =IFERROR(COUNTIF(MenuFood,">0")/COUNTA(MenuItem),0)
Commission in control   =IFERROR(1-MIN(BlendedComm/CommLimit,1),0)
Direct-order mix        =IFERROR(MIN(DirectShare/DirectGoal,1),0)
Gross margin            =IFERROR(MIN((1-FoodCostPct)/0.7,1),0)
Kitchen Score           =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `INDEX`, `MATCH`, `COUNTIF`, `COUNTA`, `AVERAGE`, `MAX`, `MIN`,
`IFERROR`, color scales (net %, kitchen score), data bars (brand share) and
conditional formatting (high commission, low inventory).

---

## 6. Printables

The 12-page PDF is print-ready as-is (US Letter) — an item-margin card, platform
P&L, packaging sheet, prep list & a promotions tracker for the line. Print any
tab: File ▸ Print ▸ fit to width.

> A business tool, not financial or accounting advice — confirm figures with your
> own books.

---

## 7. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
