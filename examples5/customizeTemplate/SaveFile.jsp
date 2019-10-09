<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*"%>
<%
// This page is used to receive the file submitted by AceoffixCtrl.
FileRequest freq = new FileRequest(request, response);
freq.saveToFile(request.getSession().getServletContext().getRealPath("customizeTemplate/doc/") + "/" + freq.getFileName());
freq.close();
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
  <head>
    
    <title></title>
    
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
     <br>
  </body>
</html>
