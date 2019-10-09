<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%
String strMht = request.getParameter("mht");
String htmlURL = "";
if((strMht != null)&&(strMht.length() > 0)){
	Random rd = new Random();
	htmlURL = strMht + "?rd=" + String.valueOf(rd.nextInt(1000));
}
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title></title>
</head>
<body>
    <div style="width:810px; margin:0 auto;">
    <p>The following HTML page is submitted by Aceoffix. Back to the <a href="../index.jsp">Example Center</a>.</p>
    <iframe width="800" height="600" src="<%=htmlURL%>"></iframe>
    </div>
</body>
</html>
