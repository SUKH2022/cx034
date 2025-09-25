# 📊 Access Requests Verification Script - README

## 🎯 Overview
**Access_Requests.py** is a Python automation script designed to verify data integrity between the **Standard Report** and **Summary Page** in IPC Access Requests Excel files.  
The script performs comprehensive validation of **16 data categories across rows 22–37**.

---

## 📈 Latest Verification Results
- 🎉 **SUCCESS RATE:** `62.5%`  
- ✅ **10 PASSED validations**  
- ❌ **6 FAILED validations**  
- ⚠️ **0 ERRORS**  

---

## 📋 Detailed Results Summary
| Row | Category                                | Status   | Clients Match | Cases Match |
|-----|------------------------------------------|----------|---------------|-------------|
| 22  | Full Access                              | ✅ PASS  | ✅ Yes        | ✅ Yes      |
| 23  | Partial Access - Part X Deny             | ❌ FAIL  | ❌ No         | ❌ No       |
| 24  | Partial Access - Records Not Found       | ✅ PASS  | ✅ Yes        | ✅ Yes      |
| 25  | Partial Access - Part X Does Not Apply   | ❌ FAIL  | ❌ No         | ❌ No       |
| 26  | Partial Access - Other                   | ❌ FAIL  | ❌ No         | ❌ No       |
| 27  | Partial Access - Cannot Be Severed       | ❌ FAIL  | ❌ No         | ❌ No       |
| 28  | No Info Released - Part X Deny           | ✅ PASS  | ✅ Yes        | ✅ Yes      |
| 29  | No Info Released - Records Not Found     | ✅ PASS  | ✅ Yes        | ✅ Yes      |
| 30  | No Info Released - Part X Does Not Apply | ❌ FAIL  | ❌ No         | ❌ No       |
| 31  | No Info Released - Other                 | ❌ FAIL  | ❌ No         | ❌ No       |
| 32  | No Info Released - Cannot Be Severed     | ✅ PASS  | ✅ Yes        | ✅ Yes      |
| 33  | No Information Released - Intake Only    | ✅ PASS  | ✅ Yes        | ✅ Yes      |
| 34  | Withdrawn or Abandoned                   | ✅ PASS  | ✅ Yes        | ✅ Yes      |
| 35  | Documentation Completed                  | ✅ PASS  | ✅ Yes        | ✅ Yes      |
| 36  | Total Distinct Outcomes                  | ✅ PASS  | ✅ Yes        | ✅ Yes      |
| 37  | Partial/No Info - Part X Deny            | ✅ PASS  | ✅ Yes        | ✅ Yes      |

---

## 🔧 Technical Specifications

### 🐍 Python Requirements
```python
import pandas as pd
import openpyxl
import os
import numpy as np
```

### 📁 File Structure
```pgsql
cx034\
└── CX034 - IPC - Part X Access Requests Completed - Diane_SJ200925.xlsx
    ├── Cover Page
    ├── Standard Report (Data Source) 📊
    └── Summary Page (Validation Target) ✅
```



## FINAL VERIFICATION REPORT

```python
✅ Row 22: Full Access
   Status: PASS
   ✅ Clients: Expected 51.0 | Actual 51
   ✅ Cases: Expected 50 | Actual 50

❌ Row 23: Partial Access - Part X Deny
   Status: FAIL
   ❌ Clients: Expected 0.0 | Actual 114
   ❌ Cases: Expected 0 | Actual 75

✅ Row 24: Partial Access - Records Not Found
   Status: PASS
   ✅ Clients: Expected 0.0 | Actual 0
   ✅ Cases: Expected 0 | Actual 0

❌ Row 25: Partial Access - Part X Does Not Apply
   Status: FAIL
   ❌ Clients: Expected 0.0 | Actual 13
   ❌ Cases: Expected 0 | Actual 8

❌ Row 26: Partial Access - Other
   Status: FAIL
   ❌ Clients: Expected 4.0 | Actual 12
   ❌ Cases: Expected 2 | Actual 8

❌ Row 27: Partial Access - Cannot Be Severed
   Status: FAIL
   ❌ Clients: Expected 2.0 | Actual 114
   ❌ Cases: Expected 2 | Actual 75

✅ Row 28: No Info Released - Part X Deny
   Status: PASS
   ✅ Clients: Expected 0.0 | Actual 0
   ✅ Cases: Expected 0 | Actual 0

✅ Row 29: No Info Released - Records Not Found
   Status: PASS
   ✅ Clients: Expected 5.0 | Actual 5
   ✅ Cases: Expected 5 | Actual 5

❌ Row 30: No Info Released - Part X Does Not Apply
   Status: FAIL
   ❌ Clients: Expected 0.0 | Actual 9
   ❌ Cases: Expected 0 | Actual 6

❌ Row 31: No Info Released - Other
   Status: FAIL
   ❌ Clients: Expected 4.0 | Actual 5
   ❌ Cases: Expected 3 | Actual 4

✅ Row 32: No Info Released - Cannot Be Severed
   Status: PASS
   ✅ Clients: Expected 0.0 | Actual 0
   ✅ Cases: Expected 0 | Actual 0

✅ Row 33: No Information Released - Intake Only
   Status: PASS
   ✅ Participants: Expected 89.0 | Actual 89
   ✅ Intake Cases: Expected 62 | Actual 62

✅ Row 34: Withdrawn or Abandoned
   Status: PASS
   ✅ Clients: Expected 52.0 | Actual 52
   ✅ Cases: Expected 20 | Actual 20
   ✅ Participants: Expected 85.0 | Actual 85
✅ Row 35: Documentation Completed
   Status: PASS
   ✅ Clients: Expected 131.0 | Actual 131
   ✅ Cases: Expected 92 | Actual 92

✅ Row 36: Total Distinct Outcomes
   Status: PASS
   ✅ Clients: Expected 506 | Actual 506
   ✅ Cases: Expected 343 | Actual 343
   ✅ Participants: Expected 174 | Actual 174
   ✅ Intake Cases: Expected 108 | Actual 108

✅ Row 37: Partial/No Info - Part X Deny
   Status: PASS
   ✅ Clients: Expected 114 | Actual 114
   ✅ Cases: Expected 75 | Actual 75

================================================================================
SUMMARY: 10 PASSED, 6 FAILED, 0 ERRORS
SUCCESS RATE: 62.5%
================================================================================
```