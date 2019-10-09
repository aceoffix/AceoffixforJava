<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*, com.acesoft.aceoffix.wordreader.*"%>
<%
// This page is used to receive the data submitted by AceoffixCtrl.
WordDocument wd = new WordDocument(request, response);
StringBuilder sb = new StringBuilder();
sb.append("<p>EmployeeName: <span class='TypeValue'>" + wd.openDataRegion("EmployeeName").getValue() + "</span></p>");
sb.append("<p>EmployeeNumber: <span class='TypeValue'>" + wd.openDataRegion("EmployeeNumber").getValue() + "</span></p>");
sb.append("<p>Department: <span class='TypeValue'>" + wd.openDataRegion("Department").getValue() + "</span></p>");
sb.append("<p>Manager: <span class='TypeValue'>" + wd.openDataRegion("Manager").getValue() + "</span></p>");
sb.append("<p>FromDate: <span class='TypeValue'>" + wd.openDataRegion("FromDate").getValue() + "</span></p>");
sb.append("<p>ToDate: <span class='TypeValue'>" + wd.openDataRegion("ToDate").getValue() + "</span></p>");
sb.append("<p>Reason: <span class='TypeValue'>" + wd.openDataRegion("Reason").getValue() + "</span></p>");

// In general, SaveDataPage do not need to dispaly anything. But if you call ShowPage method, this page will be shown in a popup dialog box after saving document. 
wd.showPage(800, 600);
// Once you saved data, please call this method finally.
wd.close();
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Extracts data from the Word document</title>
    <style type="text/css">
        .TypeValue{ color:#0000FF; text-decoration:underline;}
    </style>
</head>
<body>
    <div>
        <p><img src="../image/word.png" alt="Word Document" style="float:left;" /> 
        The following data was just typed by you and was submitted by AceoffixCtrl when you clicked the save button. This demo shows how to get the data inputted by user from Word document. 
        Typically, you can save the following data to database.</p>
        <%=sb.toString()%>
    </div>
</body>
</html>
