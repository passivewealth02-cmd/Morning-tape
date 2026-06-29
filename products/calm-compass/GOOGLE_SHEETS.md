# Calm Compass™ — Google Sheets Edition

Production-ready Google Sheets formulas mirroring the 15-sheet Excel
workbook. Same tab order: **Welcome, Dashboard, Check-In, Habits, Daily
Planner, Social Prep, Reflection, Goals, Self-Care, Sleep, Exercise,
Gratitude, Resources, Progress, Settings**.

> Build **Settings** first — it defines the dropdown lists. Then define
> the cross-sheet named ranges below.

> ⚠️ Keep the **Welcome-tab disclaimer** in the Google Sheets version too
> (wellness tool, not medical advice).

---

## 1. Settings — named ranges

Controls: `UserName` (C6), `WeekStart` (C7), `HabitGoal` (C8),
`SleepTarget` (C9), `MindGoal` (C10).

Dropdown lists (Data → Named ranges):
`ScaleList, YesNoList, GoalCatList, StatusList, IntensityList,
QualityList, SelfCareList, ResTypeList`.

---

## 2. Cross-sheet named ranges

| Range | Points to |
| ----- | --------- |
| `CheckDate` | `'Check-In'!A5:A64` |
| `CheckMood` | `'Check-In'!B5:B64` |
| `CheckEnergy` | `'Check-In'!C5:C64` |
| `CheckStress` | `'Check-In'!D5:D64` |
| `CheckSleep` | `'Check-In'!E5:E64` |
| `CheckExercise` | `'Check-In'!G5:G64` |
| `CheckMindful` | `'Check-In'!H5:H64` |
| `CheckJournaled` | `'Check-In'!I5:I64` |
| `JournalStreak` | `'Check-In'!L5:L64` (helper) |
| `HabitDone` | `Habits!K5:K64` |
| `HabitPct` | `Habits!L5:L64` |
| `HabitStreak` | `Habits!M5:M64` |
| `GoalStatus` | `Goals!E5:E64` |

---

## 3. Daily Check-In

Columns: `A Date · B Mood · C Energy · D Stress · E Sleep Hrs · F Water ·
G Exercised · H Mindful Min · I Journaled · J Notes · L Streak (helper)`.

Journaling-streak helper (col L, fill down):

```sheets
L5  =IF($A5="","",IF($I5="Yes",1,0))
L6  =IF($A6="","",IF($I6="Yes",N(L5)+1,0))   ' drag down to L64
```

Validation: `B5:D64` whole number 1–10; `G5:G64` and `I5:I64` → `YesNoList`.
Mood color scale (B): 1=soft red → 5=gold → 10=mint. Stress (D): reversed.

---

## 4. Habits

Columns: `A Date`, `B:J` the 9 habits (Yes/blank), `K Done`, `L Daily %`,
`M Streak (helper)`.

```sheets
K5 =IF($A5="","",COUNTIF(B5:J5,"Yes"))
L5 =IF($A5="","",K5/9)
M5 =IF($A5="","",IF(K5>=HabitGoal,1,0))
M6 =IF($A6="","",IF(K6>=HabitGoal,N(M5)+1,0))   ' drag down
```

Habit summary block (rows 69–77): per-habit `=COUNTIF(col5:col64,"Yes")`
and consistency `=IFERROR(times/MAX(COUNT(CheckDate),1),0)`.

Conditional formatting: "Yes" cells → mint; Daily % → data bar.

---

## 5. Goals · Sleep · the simple logs

```sheets
Goals!D (Progress)  -> percent, data bar; Status validated to StatusList
Sleep!D (Total)     -> color scale 4→7→9 hrs
```

`GoalStatus → Goals!E5:E64` for the dashboard "Goals Completed" count.

---

## 6. Dashboard KPIs

```sheets
Avg Mood          =IFERROR(AVERAGE(CheckMood),0)
Habits Completed  =SUM(HabitDone)
Routine Score     =IFERROR(AVERAGE(HabitPct),0)            [0%]
Avg Sleep         =IFERROR(AVERAGE(CheckSleep),0)
Exercise Sessions =COUNTIF(CheckExercise,"Yes")
Journaling Streak =IFERROR(LOOKUP(2,1/(JournalStreak<>""),JournalStreak),0)
Mindful Minutes   =SUM(CheckMindful)
Goals Completed   =COUNTIF(GoalStatus,"Complete")
```

Charts: Mood trend (line, `CheckMood` over `CheckDate`), Energy trend,
Sleep trend, and Habit Consistency (bar, from the Habits summary block).

---

## 7. Progress — Wellness Score

```sheets
=IFERROR(
   0.30*(AVERAGE(CheckMood)/10)
 + 0.30*AVERAGE(HabitPct)
 + 0.20*MIN(AVERAGE(CheckSleep)/SleepTarget,1)
 + 0.20*(COUNTIF(CheckJournaled,"Yes")/MAX(COUNT(CheckDate),1)),
 0)
```

Format as a big percentage. It's a gentle blend — a guide, not a grade.

---

## 8. Apps Script (optional) — daily reminder

```javascript
function dailyNudge() {
  // Optional: email yourself a gentle check-in reminder.
  MailApp.sendEmail(Session.getActiveUser().getEmail(),
    "Calm Compass — daily check-in 🌿",
    "Take 30 seconds for today's check-in. One calm day at a time.");
}
function installTrigger() {
  ScriptApp.newTrigger('dailyNudge').timeBased().everyDays(1).atHour(20).create();
}
```

---

## 9. Brand palette

| Token | Hex |
| ----- | --- |
| Primary | `#1B4F48` |
| Accent (Gold) | `#937356` |
| Gold Light | `#C9A86A` |
| Surface | `#E5D3BA` |
| Mint | `#75E6C1` |
| Ivory | `#FBF8F2` |
| Warning (disclaimer) | `#FBF0E2` |

Two-row gold-divider headers, gold-topped KPI cards, soft mood
color-scales. Keep it calm, spacious, and uncluttered.
