<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*"%>
<%
// This page is used to receive the file submitted by AceoffixCtrl.
FileRequest freq = new FileRequest(request, response);
freq.saveToFile(request.getSession().getServletContext().getRealPath("wordParagraph/doc/") + "/" + freq.getFileName());
freq.close();
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title></title>
</head>
<body>
    <div>
    
    </div>
</body>
</html>
