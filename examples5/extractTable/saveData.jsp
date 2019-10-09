<%@ page language="java"
	import="java.util.*,com.acesoft.aceoffix.*, java.awt.*, com.acesoft.aceoffix.wordreader.*"
	pageEncoding="UTF-8"%>
<%
	
        WordDocument doc = new WordDocument(request,response);
        DataRegion dataReg = doc.openDataRegion("table");
        Table table = dataReg.openTable(1);
     
        out.print("Table Data:<br/><br/>");
        StringBuilder dataStr = new StringBuilder();
        for (int i = 1; i <= table.getRowsCount(); i++)
        {
            dataStr.append("<div style='width:220px;'>");
            for (int j = 1; j <= table.getColumnsCount(); j++)
            {
                dataStr.append("<div style='float:left;width:70px;border:1px solid red;'>"+table.openCellRC(i,j).getValue()+"</div>");
            }
            dataStr.append("</div>");
        }
        out.print(dataStr.toString());
   
		doc.showPage(350, 300);
		doc.close();
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
	<head>

		<title>My JSP 'SaveFile.jsp' starting page</title>

		<meta http-equiv="pragma" content="no-cache">
		<meta http-equiv="cache-control" content="no-cache">
		<meta http-equiv="expires" content="0">
		<meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
		<meta http-equiv="description" content="This is my page">
		<!--
	<link rel="stylesheet" type="text/css" href="styles.css">
	-->

	</head>

	<body>
	</body>
</html>
