---
title: Workbook.RefreshAll method (Excel)
keywords: vbaxl10.chm199135
f1_keywords:
- vbaxl10.chm199135
api_name:
- Excel.Workbook.RefreshAll
ms.assetid: c1a956dc-263c-5c24-3b51-fc4af22dcd33
ms.date: 10/06/2026
ms.localizationpriority: medium
---


# Workbook.RefreshAll method (Excel)

Refreshes all external data ranges and PivotTable reports in the specified workbook.


## Syntax

_expression_.**RefreshAll**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

Objects that have the **[BackgroundQuery](Excel.PivotCache.BackgroundQuery.md)** property set to **True** are refreshed in the background.

If a Power Query query that loads data into a worksheet table returns an error when it's refreshed, **RefreshAll** doesn't raise a run-time error, even if background refresh is turned off for every connection. The other connections are still refreshed, and that table keeps its previous data. To detect these errors, refresh each table's query separately by calling the **[QueryTable.Refresh](Excel.QueryTable.Refresh.md)** method with the _BackgroundQuery_ argument set to **False**, and handle the error that it raises. To get the query table of a table that Power Query loads, use the **[ListObject.QueryTable](Excel.ListObject.QueryTable.md)** property.


## Example

This example refreshes all external data ranges and PivotTable reports in the third workbook.

```vb
Workbooks(3).RefreshAll
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
