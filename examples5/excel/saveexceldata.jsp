<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*, com.acesoft.aceoffix.excelreader.*"%>
<%
// This page is used to receive the data submitted by AceoffixCtrl.
Workbook wb = new Workbook(request, response);
Sheet sheet1 = wb.openSheet("Purchase Order");

StringBuilder sb = new StringBuilder();
sb.append("<p>OrderNumber: <span class='TypeValue'>" + sheet1.openCell("D8").getValue() + "</span></p>");

Table table1 = sheet1.openTable("B18:J23");
sb.append("<p>ProductItems: <br><table border=1><tr><td>QTY</td><td>DESCRIPTION</td><td>UNIT PRICE</td><td>AMOUNT</td></tr>");
while (!table1.getEOF()){
    String strCells = "";
    if (!table1.getDataFields().getIsEmpty()){
        strCells = strCells + "<td>" + table1.getDataFields().get(0).getValue() + "</td>";
        strCells = strCells + "<td>" + table1.getDataFields().get(1).getValue() + "</td>";
        strCells = strCells + "<td>" + table1.getDataFields().get(7).getValue() + "</td>"; // Skip the merged columns.
        strCells = strCells + "<td>" + table1.getDataFields().get(8).getValue() + "</td>";
        sb.append("<tr>" + strCells + "</tr>");
    }
    table1.nextRow();
}
table1.close();
sb.append("</table></p>");

sb.append("<p>Subtotal: <span class='TypeValue'>" + sheet1.openCell("J24").getValue() + "</span></p>");
sb.append("<p>Shipping: <span class='TypeValue'>" + sheet1.openCell("J25").getValue() + "</span></p>");
sb.append("<p>Tax: <span class='TypeValue'>" + sheet1.openCell("J26").getValue() + "</span></p>");
sb.append("<p>Other: <span class='TypeValue'>" + sheet1.openCell("J27").getValue() + "</span></p>");
sb.append("<p>Total: <span class='TypeValue'>" + sheet1.openCell("J28").getValue() + "</span></p>");

// In general, SaveDataPage do not need to dispaly anything. But if you call ShowPage method, this page will be shown in a popup dialog box after saving document. 
wb.showPage(800, 600);
// Once you saved data, please call this method finally.
wb.close();
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Extracts data from the Excel spreadsheet</title>
    <style type="text/css">
        .TypeValue{ color:#0000FF; text-decoration:underline;}
    </style>
</head>
<body>
    <div>
        <p><img src="../image/excel.png" alt="Excel Spreadsheet" style="float:left;" /> 
        The following data was just submitted by AceoffixCtrl when you clicked the save button. This demo shows how to get the data inputted by user and other data from Excel spreadsheet. 
        Typically, you can save the following data to database.</p>
        <%=sb.toString()%>
    </div>
</body>
</html>
