# 📊 Multi File Merge Tool (No More VLOOKUP)

## What this is

This project is mainly to avoid doing multiple VLOOKUPs in Excel.

Normally we have 2–3 files and we keep doing VLOOKUP again and again to match data.
That is time taking and also chances of mistake.

So this script does same thing in one go using Python.

---

## What it does

* Takes 3 files:

  * main quality report
  * internal / edge report
  * client survey data

* User can:

  * select which columns they want
  * select which columns to match

* Then code:

  * merges data step by step
  * gives one final combined file

---

## Simple flow

1. Load all files
2. Select columns from merge report
3. Select columns from internal file
4. First merge happens using **Token**
5. Drop unwanted columns from client file
6. Select matching columns between files
7. Final merge
8. Excel output generated

---

## Why I made this

* Too much manual VLOOKUP work
* Same thing repeating again and again
* Time waste + error chances
* Needed something flexible (not fixed columns)

---

## Features

* No hardcoding → user selects columns
* Works with:

  * single (5)
  * multiple (1,3,6)
  * range (2-10)
* Basic validation added
* Removes:

  * empty columns
  * duplicate columns (_y after merge)

---

## Input requirement

* Files should have proper headers
* Common column like **Token** is required
* Column numbering starts from 1 (not 0)

---

## Output files

* `prelim_merge_output.xlsx` → first merge
* `final_output.xlsx` → final result

---

## How to run

```bash
python your_script_name.py
```

Then just follow prompts and enter column numbers.

---

## Things to remember

* Enter correct column numbers
* If wrong input once, just re-enter
* Large files → may take time

---

## Limitations

* Fully manual input (no UI)
* Depends on user selection
* No auto column matching yet

---

## Future plan

* Add simple UI (maybe Streamlit)
* Auto match columns
* Reduce manual steps
* Better error messages

---

## Author note

Made this mainly for daily work to save time from Excel VLOOKUP.

Not perfect code, but works well for real use.

---

## If you are using this

If you also deal with multiple reports and merging, this should help you save good time.
