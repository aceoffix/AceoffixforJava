<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*"%>
<%
FileRequest freq = new FileRequest(request, response);
freq.saveToFile(request.getSession().getServletContext().getRealPath("applyTemplate/doc/") + "/" + freq.getFileName());
freq.close();
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title></title>
</head>
<body>
    <div>
    
    </div>
</body>
</html>
