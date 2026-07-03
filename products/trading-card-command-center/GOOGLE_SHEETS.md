# Trading Card Collection Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Collection, Purchases, Sales, Trades, Grading,
Value Analytics, Wishlist, Duplicates, Deck Builder, Card Vault, Analytics,
Settings**.

> Build **Settings** first (collection details + dropdown lists), then the
> Master Collection, then the trackers, then the Dashboard. Add the named
> ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `CollName` (C6), `Collector` (C7), `PrimaryGame` (C8),
`HomeCurr` (C9), `ValueGoal` (C10), `GrowthTarget` (C11), `CardBudget` (C12).

Lists: `GameList, RarityList, LangList, CondList, LocList, GraderList,
MarketList, CurrencyList, PriorityList, PlanList, GradeStatusList, YesNoList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `ColName` | `Collection!A5:A44` | `GradeStatus` | `Grading!E5:E34` |
| `ColSet` | `Collection!B5:B44` | `GradeAdded` | `Grading!J5:J34` |
| `ColRarity` | `Collection!E5:E44` | `TradePartner` | `Trades!A5:A34` |
| `ColQty` | `Collection!G5:G44` | `TradeNet` | `Trades!G5:G34` |
| `ColLocation` | `Collection!H5:H44` | `SaleNet` | `Sales!H5:H34` |
| `ColPaid` | `Collection!J5:J44` | `WishCard` | `Wishlist!A5:A34` |
| `ColEach` | `Collection!K5:K44` | `WishBought` | `Wishlist!F5:F34` |
| `ColTotal` | `Collection!L5:L44` | `DupCard` / `DupPlan` | `Duplicates!A/E` |
| `ColPhoto` | `Collection!M5:M44` | `DeckWins` / `DeckLosses` | `'Deck Builder'!H6/H7` |
| `PurchTotal` | `Purchases!H5:H44` | `HealthRange` | `Analytics!C7:C12` |

---

## 3. Master Collection — live value per card

```sheets
Total Value (row)  =IF(A5="","",G5*K5)          (qty × est. value)
Row P/L            =L5-J5
```

**📷 Card photos:** the `Photo?` column (M) tracks coverage. In Google Sheets
you can go further — add an *Image* column and use
`=IMAGE("https://…/your-card.jpg")` to show a thumbnail **in the cell**, or
Insert ▸ Image ▸ *Insert image in cell*. In the **Card Vault** tab, click a
photo box and Insert ▸ Image (Excel: Insert ▸ Pictures ▸ This Device).

---

## 4. Dashboard — the 12 KPIs

```sheets
Total Cards        =SUM(ColQty)
Collection Value   =SUM(ColTotal)
Purchase Cost      =SUM(ColPaid)
Profit / Loss      =SUM(ColTotal)-SUM(ColPaid)
Growth             =IFERROR(SUM(ColTotal)/SUM(ColPaid)-1,0)
Top Card Value     =MAX(ColEach)
Photos on File     =IFERROR(COUNTIF(ColPhoto,"Yes")/MAX(COUNTA(ColName),1),0)
Cards Graded       =COUNTIF(GradeStatus,"Returned")
Awaiting Grading   =COUNTIF(GradeStatus,"Pending")
Trades Completed   =COUNTA(TradePartner)
Wishlist Progress  =IFERROR(COUNTIF(WishBought,"Yes")/MAX(COUNTA(WishCard),1),0)
Collection Health  =IFERROR(AVERAGE(HealthRange),0)
```

Charts: Value by Rarity (donut, via `SUMIF(ColRarity,…,ColTotal)` summary),
Value by Set (column), Collection Value Over Time (line), Grading Before vs
After (column). Turn off auto data labels.

---

## 5. Analytics — Collection Health Score

```sheets
Value growth vs target  =IFERROR(MIN((SUM(ColTotal)/SUM(ColPaid))/GrowthTarget,1),0)
Grading returned        =IFERROR(COUNTIF(GradeStatus,"Returned")/MAX(COUNTA(GradeStatus),1),0)
Wishlist progress       =IFERROR(COUNTIF(WishBought,"Yes")/MAX(COUNTA(WishCard),1),0)
Storage documented      =IFERROR(SUMPRODUCT((ColLocation<>"")*(ColName<>""))/MAX(COUNTA(ColName),1),0)
Duplicates planned      =IFERROR(SUMPRODUCT((DupPlan<>"")*(DupCard<>""))/MAX(COUNTA(DupCard),1),0)
Tournament win rate     =IFERROR(DeckWins/MAX(DeckWins+DeckLosses,1),0)
Health Score            =IFERROR(AVERAGE(C7:C12),0)
```

Power features: `ARRAYFORMULA`, `QUERY` ("top 5 by value":
`=QUERY(Collection!A5:L,"select A,L order by L desc limit 5",0)`),
`FILTER`/`SORT`/`UNIQUE`, all wrapped in `IFERROR`.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, value data-bars, mint
"owned/returned" flags. Keep it premium and consistent — that polish is what
makes it feel like software, not a spreadsheet.
