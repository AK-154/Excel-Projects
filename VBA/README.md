# OOH Advertising Cost Tracker

## Overview
This project is an Excel workbook that tracks out-of-home (OOH) advertising placement costs for brands like Douze 2, Chapil late, and El mond across various Nigerian states. Each brand team researched costs independently, and I used VBA to restrict sheet access—preventing copy-pasting between teams while allowing senior management to view all data for average calculations.

## Key Features
- **Data**: Covers brands, states, locations, media owners, board types, and monthly rates (e.g., Douze 2: ₦27.2M, Chapil late: ₦13.2M, El mond: ₦58.7M).
- **VBA Access Control**:
  - Password `1678`: Douze 2 team sees only their sheet.
  - Password `2467`: Chapil late team sees only their sheet.
  - Password `4782`: El mond team sees only their sheet.
  - Password `3509`: Senior management sees all sheets.
- **Why One Workbook?**: Prevents data copying, centralizes averages, simplifies updates, and reduces leakage risk.

## Tools Used
- **Excel**: For data storage and calculations.
- **VBA**: To hide sheets and control access via passwords.

## How to Use
1. Download `OOH_Cost_Tracker.xlsm` from this repository.
2. Open it in Excel (enable macros when prompted).
3. Enter your password in the login form:
   - Team passwords: `1678`, `2467`, or `4782` for specific sheets.
   - Management password: `3509` to see all.
4. Explore your assigned sheet or the full dataset (if authorized).

## Challenges and Learnings
- **Challenge**: Ensuring teams couldn’t copy each other’s data.
- **Solution**: Used VBA to hide sheets and enforce password access.
- **Learned**: VBA’s power for data security, though it’s not unbreakable.

## Why Not Separate Workbooks?
- **Copy-Paste Risk**: Teams could share files and copy data.
- **Management Effort**: Harder to compare averages across files.
- **Sync Issues**: Updates wouldn’t reflect universally.

## Critical Notes
- **VBA Limits**: Savvy users could bypass it with developer tools—Excel needs better native controls.
- **Scalability**: Works now, but a database might be better as brands grow.

## Get the Code
- Check the VBA scripts in the repository:
  - `Workbook_Open`: Hides sheets on startup.
  - `LoginButton_Click`: Controls sheet visibility by password.

## Next Steps
- Add more brands as data grows.
- Explore a database for long-term scalability.

---

![image](https://github.com/user-attachments/assets/05a0502f-2b72-431a-bd68-aff713f07970)



Questions? Reach out or fork this repo to adapt it!

