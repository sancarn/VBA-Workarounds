# Excel crash bug - corrupt chart object

This document is to demonstrate a nasty Microsoft Excel bug. The bug occurs while VBA opens a corrupt Excel workbook, and checks for vertical page breaks. The bug causes Excel to crash, and unfortunately one is unable to test for the corrupt behaviour with the core Excel libraries. Additionally trying to open with `CorruptLoad.repair` doesn't pick up or fix the issue.

## Version information:

```
Microsoft Excel for Microsoft 365 MSO (16.0.13328.20334) 32-bit
Version 2010 (Build 13328.20408)
```

## Code executed in VBA

```vb    
Sub tt()
    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=ThisWorkbook.Path & "\FailsOnThisWorkbook.xlsx", UpdateLinks:=False, CorruptLoad:=Excel.XlCorruptLoad.xlRepairFile)
    Dim ws As Worksheet: Set ws = wb.Sheets(1)
    
    'Crash occurrs here:
    Debug.Print ws.VPageBreaks.Count
    wb.Close False
End Sub
```

In general it seems  the crash occurs in that general vicinity. Sometimes it doesn't actually crash until close statement, other times it crashes immediately. Often it will at least report the following VBA error:

![img](https://github.com/sancarn/VBA-Workarounds/blob/main/CorruptWorkbook-CrashOn-VPageBreaks-Count/VBAError.png?raw=true)

And either trying to debug or trying to end at this point will cause a crash.

![img2](https://github.com/sancarn/VBA-Workarounds/blob/main/CorruptWorkbook-CrashOn-VPageBreaks-Count/VBACrash.PNG?raw=true)

## Investigations

Initially when this bug was shown to me I had a lot of struggle isolating the bug, as with many bugs in Excel where the application just crashes. After a bit of time I isolated the bug to the workbook provided in this repository. Then we could begin assessing the issues.

> Note: The test workbook which isolates the issue has gone through several versions. The chart was still crashing when it contained data, but the data was removed in order to make the XML neater.


### 1. What happens when you delete the bugged out chart `Chart 6`?

No surprise, when the chart is removed it fixes the issue. So we can at least say that this problem is due to the chart's existence.

What are the differences?

#### Data present in workbook with buggy chart, but missing from the "fixed" workbook:

##### In `xl\worksheets\sheet1.xml`

```
<drawing r:id="rId2"/>
```

##### In `xl\worksheets\_rels\sheet1.xml.rels`

```
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
```

This is to be expected and relates to drawing1.xml

##### The entirity of drawing1.xml and drawing2.xml

Curiously `drawing2.xml` doesn't appear to be referenced anywhere. `drawing2.xml.rels` also don't exist. This leads me to wonder whether this is actually what causes the bug.




### 2. 



