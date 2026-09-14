# IB Excel Dojo

Practice workbook for investment-banking Excel.

Not a course. A gym: shortcuts, format rehab, formula drills, a balancing mini 3-statement, comps, a sensitivity grid, and a check dashboard.

**Color language:** blue font = input · black = formula · green = cross-sheet link · yellow fill = toggle me.

Hypothetical numbers for a fictional company (Apex Industrial). Not investment advice. Not a live deal model.

## File map

| Path | What |
|---|---|
| `IB_Excel_Dojo.xlsx` | The 9-tab workbook |
| `index.html` | One-page landing / GitHub Pages home |
| `src/build_ib_excel_dojo.py` | Script that generated the xlsx |
| `LICENSE.txt` | Share terms + disclaimer |

## Use the workbook

1. Open `00_Start` once
2. Drill starred Week 1 rows on `01_Shortcuts`
3. Fix `02_Format_Gym` with no mouse
4. Yellow cells on `03_Formula_Dojo` and `04_3Statement`
5. Tick days on `08_Plan`

## Push this repo from Windows

You do **not** need a Mac. You do need Git installed once.

### 1. Install Git (one time)

PowerShell (Admin not required):

```powershell
winget install --id Git.Git -e --source winget
```

Close and reopen the terminal after it finishes. Check:

```powershell
git --version
```

Optional but nice — GitHub CLI, so you can create the repo from the terminal instead of the website:

```powershell
winget install --id GitHub.cli -e --source winget
gh --version
gh auth login
```

(`gh auth login` → GitHub.com → HTTPS → login in the browser.)

### 2. Unzip this folder

Example: `Downloads\ib-excel-dojo-repo`

### 3. First push

```powershell
cd $HOME\Downloads\ib-excel-dojo-repo
git init
git add .
git status
git commit -m "Initial commit: IB Excel Dojo v1"
git branch -M main
```

If you installed `gh`:

```powershell
gh repo create ib-excel-dojo --public --source=. --remote=origin --push
```

If you did **not** install `gh`, create an empty public repo on github.com named `ib-excel-dojo` (no README), then:

```powershell
git remote add origin https://github.com/YOURUSER/ib-excel-dojo.git
git push -u origin main
```

Windows will prompt for GitHub login the first time. Use a **Personal Access Token** as the password if it asks, or pick the browser sign-in helper. Do not type your GitHub account password into Git.

### 4. GitHub Pages (the public link)

Repo → Settings → Pages → Deploy from branch → `main` / `/ (root)` → Save.

Site URL becomes `https://YOURUSER.github.io/ib-excel-dojo/`  
The Download button on `index.html` serves the xlsx.

## Later updates from the same PC

```powershell
cd $HOME\Downloads\ib-excel-dojo-repo
git add .
git commit -m "Describe what changed"
git push
```

That is the fun part. Website is just the remote. The work happens in the folder.

## Mac users

Same commands. Install Git via Xcode CLT or Homebrew (`brew install git gh`) instead of winget.
