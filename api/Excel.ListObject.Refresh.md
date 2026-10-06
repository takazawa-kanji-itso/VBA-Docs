---
title: ListObject.Refresh method (Excel)
keywords: vbaxl10.chm734075
f1_keywords:
- vbaxl10.chm734075
api_name:
- Excel.ListObject.Refresh
ms.assetid: 7827a116-0ba4-9855-e0e9-550a85d36ed3
ms.date: 10/06/2026
ms.localizationpriority: medium
---


# ListObject.Refresh method (Excel)

Refreshes the data in the list.

For a list that's linked to a SharePoint site, retrieves the current data and schema for the list from the server that is running Microsoft SharePoint Foundation. If the SharePoint site is not available, calling this method returns an error.

For a list that's loaded from an external data source, such as a Power Query query, refreshes the list from that source. If evaluating the query results in an error, the method raises the same error as the **[Refresh](Excel.QueryTable.Refresh.md)** method of the list's **[QueryTable](Excel.ListObject.QueryTable.md)** object.


## Syntax

_expression_.**Refresh**

_expression_ A variable that represents a **[ListObject](Excel.ListObject.md)** object.


## Remarks

For a list that's linked to a SharePoint site, calling the **Refresh** method does not commit changes to the list in the Excel workbook. Uncommitted changes in the list in Excel are discarded when the **Refresh** method is called.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
