<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*,java.sql.*"%>
<%
// This page is used to receive the file submitted by AceoffixCtrl.
FileRequest freq = new FileRequest(request, response);
freq.saveToFile(request.getSession().getServletContext().getRealPath("passParameters/doc/") + "/" + freq.getFileName());

int id = 0;
	String userName = "";
	int age = 0;
	String sex = "";

	if (request.getParameter("id") != null
			&& request.getParameter("id").trim().length() > 0)//Get the parameter
		id = Integer.parseInt(request.getParameter("id").trim());

	if (freq.getFormField("userName") != null
			&& freq.getFormField("userName").trim().length() > 0) {
		userName = freq.getFormField("userName");
	}//Get the form field

	if (freq.getFormField("age") != null
			&& freq.getFormField("age").trim().length() > 0) {
		age = Integer.parseInt(freq.getFormField("age"));
	}
	
	if (freq.getFormField("selSex") != null
			&& freq.getFormField("selSex").trim().length() > 0) {
		sex = freq.getFormField("selSex");
	}
	Class.forName("org.sqlite.JDBC");
	String strUrl = "jdbc:sqlite:"
			+ this.getServletContext().getRealPath("data/")
			+ "/SendParameters.db";
	Connection conn = DriverManager.getConnection(strUrl);
	Statement stmt = conn.createStatement();
	String strsql = "Update Users set UserName = '" + userName
			+ "', age = " + age + ",sex = '" + sex + "' where id = "
			+ id;
	stmt.executeUpdate(strsql);
	conn.close();
	freq.showPage(300, 200);
	freq.close();
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title></title>
</head>
<body>
    <div>
    parameters:<br />
    id:<%=id %><br />
    userName:<%=userName%><br />
    age:<%=age%><br />
    sex:<%=sex%><br />
    </div>
</body>
</html>
