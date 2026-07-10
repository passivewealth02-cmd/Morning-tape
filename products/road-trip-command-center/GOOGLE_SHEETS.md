# Road Trip Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the Excel workbook. Same tab
order: **Welcome, Dashboard, Trip Profile, Route, Itinerary, Budget, Fuel,
Vehicle, Stays, Camping, Attractions, Food, Packing, Emergency, Journal,
Gallery, Parks, Rewards, Analytics, Settings**.

> Build **Settings** first (trip + vehicle details and dropdown lists), then the
> Route, Budget, Fuel, Vehicle, Stays & Camping, then the Dashboard + Analytics.
> Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `TripName`, `StartLoc`, `Destination`, `TripStart` (date), `TripEnd`
(date), `TravelerCount` (2), `VehicleName`, `AvgMPG` (16), `FuelPrice` (3.60).

Lists: `ExpCatList, StayTypeList, ActTypeList, PackCatList, FuelTypeList,
VehTypeList, RoadCondList, ResStatusList, PriorityList, YesNoList,
VehStatusList, CurrencyList`.

---

## 2. Cross-sheet named ranges

| Range | Points to | Range | Points to |
| ----- | --------- | ----- | --------- |
| `DailyMiles` | `Route!D5:D28` | `TripBudget` | `Budget!B17` |
| `DriveTime` | `Route!E5:E28` | `SpentTotal` | `Budget!C17` |
| `TotalDistance` | `Analytics!F6` | `SavedTotal` | `Budget!H8` |
| `DriveHours` | `Analytics!F7` | `PackStatus` | `Packing!E5:E64` |
| `FuelEstimate` | `Fuel!C…` (summary) | `VehStatus` | `Vehicle!E7:E16` |
| `StayType` | `Stays!B5:B24` | `StayBooked` | `Stays!H5:H24` |
| `CampReserved` | `Camping!C5:C24` | `AttractionName` | `Attractions!A5:A34` |
| `ParkName` | `Parks!A5:A20` | `ReadinessRange` | `Analytics!C7:C10` |

---

## 3. Route & Fuel

```sheets
Total miles          =SUM(DailyMiles)
Driving hours        =SUM(DriveTime)
MPG (per fill-up)    =IFERROR(ROUND((F6-F5)/C6,1),0)   (odometer diff ÷ gallons)
Est. fuel needed     =IFERROR(TotalDistance/AvgMPG,0)
Est. fuel cost       =IFERROR(TotalDistance/AvgMPG*FuelPrice,0)
Fuel cost / mile     =IFERROR(FuelPrice/AvgMPG,0)
```

---

## 4. Dashboard — the 12 KPIs

```sheets
Days Until Departure =MAX(TripStart-TODAY(),0)
Total Distance       =TotalDistance
Driving Hours        =DriveHours
Total Budget         =TripBudget
Budget Remaining     =TripBudget-SpentTotal
Fuel Cost Estimate   =FuelEstimate
Stops Planned        =COUNTA(AttractionName)
Campgrounds Booked   =COUNTIF(CampReserved,"Yes")
Hotels Booked        =COUNTIFS(StayType,"Hotel",StayBooked,"Yes")
Packing Progress     =IFERROR(COUNTIF(PackStatus,"Yes")/COUNTA(PackStatus),0)
Vehicle Readiness    =IFERROR(COUNTIF(VehStatus,"OK")/COUNTA(VehStatus),0)
Trip Readiness       =IFERROR(AVERAGE(ReadinessRange),0)
```

Charts: Budget by Category (donut), Daily Mileage (column), Trip Readiness
(bar), Planned vs Actual (column). Turn off auto data labels.

---

## 5. Analytics — Trip Readiness Score

```sheets
Fund vs budget       =IFERROR(MIN(SavedTotal/TripBudget,1),0)
Packing complete     =IFERROR(COUNTIF(PackStatus,"Yes")/COUNTA(PackStatus),0)
Vehicle ready        =IFERROR(COUNTIF(VehStatus,"OK")/COUNTA(VehStatus),0)
Bookings confirmed   =IFERROR((COUNTIF(StayBooked,"Yes")+COUNTIF(CampReserved,"Yes"))/(COUNTA(StayBooked)+COUNTA(CampReserved)),0)
Trip Readiness Score =IFERROR(AVERAGE(C7:C10),0)
```

Power features: `ARRAYFORMULA`, `QUERY` ("spend by category", "unpacked items"),
`FILTER`/`SORT`/`UNIQUE`, all wrapped in `IFERROR`.

Photo & Memory Gallery: click a cell ▸ Insert ▸ Image ▸ **Image in cell**, or
paste `=IMAGE("your-photo-link")`.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |

Two-row gold-divider headers, gold-topped KPI cards, status color flags. Keep it
premium and consistent — that polish is what makes it feel like road trip
software, not a spreadsheet.
