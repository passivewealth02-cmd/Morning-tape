# Fitness & Meal-Prep Command Center™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the workbook. Tab order:
**Start Here, Dashboard, Goals & Stats, Meal Plan, Recipe Bank, Grocery List,
Macro Tracker, Workout Plan, Workout Log, Body Metrics, Habit Tracker,
Settings**.

> Build **Settings** and **Recipe Bank** first, then the Meal Plan, Grocery List,
> Macro Tracker, Workout Plan/Log, Body Metrics and Habit Tracker, then the
> Dashboard. Add the named ranges below (Data ▸ Named ranges).

---

## 1. Settings — controls & lists

Controls: `Name`, `StartWt` (185), `CurrentWt` (176), `GoalWt` (165),
`CalTarget` (2100), `ProteinTarget` (150), `WaterTarget` (8), `StepGoal`
(10000), `WeeklyWorkoutGoal` (5).

Lists: `SlotList, CatList, AisleList, FocusList, YesNoList`.

---

## 2. The engine tabs

```sheets
Macro Tracker    averages via =ROUND(AVERAGE(MacroCal),0) / AVERAGE(MacroProtein)
Workout Log      Volume (G) =Sets*Reps*Weight ; total =SUM(volume)
Workout Plan     Done this week =COUNTIF(WorkoutDone,"Yes")
Body Metrics     Change =C{r}-C{r-1} ; total change =C_last-C_first
Habit Tracker    averages via =ROUND(AVERAGE(HabitWater),1) etc.
Grocery List     Still to buy =COUNTIF(GroceryHave,"No")
```

Named: `RecipeName/RecipeCal/RecipeProtein`, `PlanCal`, `GroceryHave/GroceryItem`,
`MacroCal/MacroProtein`, `WorkoutDone`, `WeightVal/WeightWk`,
`HabitWater/HabitSleep/HabitSteps`.

---

## 3. Dashboard — the 12 KPIs

```sheets
Current Weight  =CurrentWt
Goal Weight     =GoalWt
Lbs to Go       =CurrentWt-GoalWt
Lost So Far     =StartWt-CurrentWt
Calorie Target  =CalTarget
Avg Calories    =IFERROR(ROUND(AVERAGE(MacroCal),0),0)
Protein Target  =ProteinTarget
Avg Protein     =IFERROR(ROUND(AVERAGE(MacroProtein),0),0)
Workouts / Wk   =COUNTIF(WorkoutDone,"Yes")&" / "&WeeklyWorkoutGoal
Steps Avg       =IFERROR(ROUND(AVERAGE(HabitSteps),0),0)
Water Avg       =IFERROR(ROUND(AVERAGE(HabitWater),1),0)
Fitness Score   =IFERROR(AVERAGE(HealthRange),0)
```

Chart: Weight Trend (line) from the Body Metrics weight column.

---

## 4. Fitness Score (6 dimensions)

```sheets
Weight to goal      =IFERROR((StartWt-CurrentWt)/(StartWt-GoalWt),0)
Protein hit         =IFERROR(MIN(AVERAGE(MacroProtein)/ProteinTarget,1),0)
Calories on target  =IFERROR(1-ABS(AVERAGE(MacroCal)-CalTarget)/CalTarget,0)
Workouts done       =IFERROR(MIN(COUNTIF(WorkoutDone,"Yes")/WeeklyWorkoutGoal,1),0)
Steps               =IFERROR(MIN(AVERAGE(HabitSteps)/StepGoal,1),0)
Water               =IFERROR(MIN(AVERAGE(HabitWater)/WaterTarget,1),0)
Fitness Score       =IFERROR(AVERAGE(C13:C18),0)
```

Power features: `AVERAGE`, `COUNTIF`, `SUM`, `MIN`, `ABS`, `ROUND`, `IFERROR`,
data bars (meal-plan calories), color scales (calories vs target, Fitness Score)
and conditional formatting (workout done / grocery have = mint).

---

## 5. Printables

The 12-page PDF is print-ready as-is (US Letter) — including a weekly meal
planner, a habit tracker grid and a weight-progress chart. Print any tab: File ▸
Print ▸ fit to width.

> A general wellness tool, not medical, nutrition or training advice — talk to a
> doctor or qualified professional before starting any new diet or program.

---

## 6. Brand palette

| Token | Hex | Token | Hex |
| ----- | --- | ----- | --- |
| Primary | `#1B4F48` | Mint | `#75E6C1` |
| Accent (Gold) | `#937356` | Ivory | `#FBF8F2` |
| Gold Light | `#C9A86A` | Surface | `#E5D3BA` |
