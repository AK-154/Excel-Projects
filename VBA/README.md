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

## VBA Code
Below are the key VBA scripts that control access to the workbook's sheets.

### Workbook_Open
This code runs when the workbook is opened. It hides all sheets except the "Intro" sheet and shows the login form:

```vba
Private Sub Workbook_Open()
    ' Hide all sheets except a default one
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name <> "Intro" Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    ' Show login form
    LoginForm.Show
End Sub

```

### LoginButton_Click
This code runs when the user clicks the login button on the form:

```vba
Private Sub Label1_Click()

End Sub

Private Sub LoginButton_Click()
Dim password As String
' Get the value from the TextBox named "password"
password = Me.password.Value
' Control sheet visibility based on password
    If password = "1678" Then
        ' Password for Douze 2: Unhide "Douze 2" and "Intro"
        Worksheets("Intro").Visible = xlSheetVisible
        Worksheets("Douze 2").Visible = xlSheetVisible
        MsgBox "Welcome! You have access to Douze 2 data."
    ElseIf password = "2467" Then
        ' Password for Chapil late: Unhide "Chapil late" and "Intro"
        Worksheets("Intro").Visible = xlSheetVisible
        Worksheets("Chapil late").Visible = xlSheetVisible
        MsgBox "Welcome! You have access to Chapil late data."
    ElseIf password = "4782" Then
        ' Password for El mond: Unhide "El mond" and "Intro"
        Worksheets("Intro").Visible = xlSheetVisible
        Worksheets("El mond").Visible = xlSheetVisible
        MsgBox "Welcome! You have access to El mond data."
    ElseIf password = "3509" Then
        ' Password to see all: Unhide all sheets
        Dim ws As Worksheet
        For Each ws In Worksheets
            ws.Visible = xlSheetVisible
        Next ws
        MsgBox "Welcome! You have access to all data."
    Else
        MsgBox "Access denied. Invalid password."
End If
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
```

## Next Steps
- Add more brands as data grows.
- Explore a database for long-term scalability.

---

![image](https://github.com/user-attachments/assets/05a0502f-2b72-431a-bd68-aff713f07970)



Questions? Reach out or fork this repo to adapt it!

