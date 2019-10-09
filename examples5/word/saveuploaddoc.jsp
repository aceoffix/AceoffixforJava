<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*"%>
<%
// This page is used to receive the file submitted by AceoffixCtrl.
FileRequest freq = new FileRequest(request, response);
freq.saveToFile(request.getSession().getServletContext().getRealPath("doc/") + "\\" + freq.getFileName());
String strDocumentText = freq.getDocumentText();
freq.showPage(500, 500);
freq.close();
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title></title>
</head>
<body>
    <div>
        <p style="background-color:Gray; color:White; font-weight:bold">The following text is the plain text of current document. Typically, you can save the following text to database.</p>
        <p><%=strDocumentText%></p>
    </div>
</body>
</html>
