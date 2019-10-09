<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*, com.acesoft.aceoffix.wordreader.*,java.io.*"%>
<%
	WordDocument doc = new WordDocument(request, response);
	byte[] bytes = null;
	String filePath = "";
	if (request.getParameter("userName") != null && request.getParameter("userName").trim().equalsIgnoreCase("Tom")) {
		bytes = doc.openDataRegion("Paragraph1").getFileBytes();
		filePath = "paragraph1.doc";
	} else {
		bytes = doc.openDataRegion("Paragraph2").getFileBytes();
		filePath = "paragraph2.doc";
	}
	doc.close();
	
	filePath = request.getSession().getServletContext().getRealPath("wordEditByUser/doc/") + "/" + filePath;
	FileOutputStream outputStream = new FileOutputStream(filePath);
	outputStream.write(bytes);
	outputStream.flush();
	outputStream.close();

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
    
</body>
</html>
