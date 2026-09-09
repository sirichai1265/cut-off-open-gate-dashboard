# Cut off / Open gate dashboard

Live: **https://sirichai1265.github.io/heung-a-cutoff/**

## Update with a new daily file

1. Drop the new `*CUT*.xls` into this folder (e.g. `9-10-CUT.xls`).
2. Right-click **`update.ps1`** → **Run with PowerShell**.
   (or run `powershell -ExecutionPolicy Bypass -File update.ps1`)

That rebuilds `index.html` + `Cut off - Open Gate Dashboard.html`, commits, and
pushes. GitHub Pages goes live ~1 minute later.

To use a specific file instead of the newest one:
`powershell -ExecutionPolicy Bypass -File update.ps1 "C:\path\to\9-10-CUT.xls"`

## Manual build (no git)

```
python build_dashboard.py 9-10-CUT.xls
```

Needs `pip install xlrd`. Uses the built-in wharf-code lookup
(`DEFAULT_WHARF_MAP` in `build_dashboard.py`) — the build prints a warning if a
new wharf code shows up that isn't in it; add it there.

## Rules

| Field            | Calculation                                        |
|------------------|----------------------------------------------------|
| Cut off (Dry)    | ETA − 24 hours                                      |
| Cut off (Reefer) | ETA − 1 hour                                        |
| 1st Return       | ETD − 5 days, counting ETD as day 1, time ignored   |

Rows with `Skip = Y` are excluded. Input columns expected (in order): Service,
Vessel, Vessel Name, Op.Liner, Seq, Vyg, Bound, Vyg Bound, Wharf, POL, POD,
Skip, No D, No L, USED, ETA Date, ETB Date, ETD Date.
