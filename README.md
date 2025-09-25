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