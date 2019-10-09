<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*"%>
<%
// This page is used to receive the file submitted by AceoffixCtrl.
String strID=request.getParameter("id");
if((strID!=null)&&(strID.length()>0)){
	FileRequest freq = new FileRequest(request, response);
	freq.saveToFile(request.getSession().getServletContext().getRealPath("AceoffixLinkParams/doc/") + "/" + freq.getFileName());
	freq.close();
	//Add code to update database by strID.
}
else
	out.println("The parameter id is invalid.");
		
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
