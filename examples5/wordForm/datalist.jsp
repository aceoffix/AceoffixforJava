<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="java.awt.*,javax.servlet.*,javax.servlet.http.*,java.sql.*,java.text.SimpleDateFormat,java.util.Date"%>

<%
	Class.forName("org.sqlite.JDBC");
	String strUrl = "jdbc:sqlite:" + this.getServletContext().getRealPath("data/") + "/demo2.db";
	Connection conn = DriverManager.getConnection(strUrl);
	Statement stmt = conn.createStatement();
	String id = request.getParameter("ID");
	ResultSet rs = stmt.executeQuery("select * from AbsenceRecord where ID = " + id);
	String docSubject = "", docName = "", docDept = "", docCause = "", docDays = "", docDate = "";
	if (rs.next()) {
		docSubject = rs.getString("Subject");
		docName = rs.getString("Name");
		docDept = rs.getString("Dept");
		docCause = rs.getString("Cause");
		docDays = rs.getString("Num");
		docDate = rs.getString("SubmitTime");
	}
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<TITLE></TITLE>
		<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
		<meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" Content="C#">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema"
			content="http://schemas.microsoft.com/intellisense/ie5">
		<link rel="stylesheet" href="images/csstg.css" type="text/css"></link>
		<link rel="stylesheet" href="css/style.css" type="text/css"></link>
	</HEAD>
	<body>
		<form id="Form1" method="post">
		</form>

		<table width='98%' class="zz-talbe"
			style="margin-top: 20px; width: 460px;border: 1px solid #999999;">
			<thead>
				<tr>
					<th width='20%' height='26' valign="middle" >
					Database field
					</th>
					<th width='80%' height='23' valign="middle">
						Database values
					</th>
				</tr>
			</thead>
			<tr>
				<td width='20%' height='26' valign="middle"
					style="background-color: #D7FFEE;border: 1px solid #999999;">
					Theme
				</td>
				<td width='80%' height='23' valign="middle" style="border: 1px solid #999999;">
					&nbsp;&nbsp;&nbsp;&nbsp;<%=docSubject%></td>
			</tr>
			<tr>
				<td width='20%' height='26' valign="middle"
					style="background-color: #D7FFEE;border: 1px solid #999999;">
					Name
				</td>
				<td width='80%' height='23' valign="middle" style="border: 1px solid #999999;">
					&nbsp;&nbsp;&nbsp;&nbsp;<%=docName%></td>
			</tr>
			<tr>
				<td width='20%' height='26' valign="middle"
					style="border: 1px solid #999999;background-color: #D7FFEE">
					Department
				</td>
				<td width='80%' height='23' valign="middle" style="border: 1px solid #999999;">
					&nbsp;&nbsp;&nbsp;&nbsp;<%=docDept%></td>
			</tr>
			<tr>
				<td width='20%' height='26' valign="middle"
					style="border: 1px solid #999999;background-color: #D7FFEE">
					Reason
				</td>
				<td width='80%' height='23' valign="middle" style="border: 1px solid #999999;">
					&nbsp;&nbsp;&nbsp;&nbsp;<%=docCause%></td>
			</tr>
			<tr>
				<td width='20%' height='26' valign="middle"
					style="border: 1px solid #999999;background-color: #D7FFEE">
					Days
				</td>
				<td width='80%' height='23' valign="middle" style="border: 1px solid #999999;">
					&nbsp;&nbsp;&nbsp;&nbsp;<%=docDays%></td>
			</tr>
			<tr>
				<td width='20%' height='26' valign="middle"
					style="border: 1px solid #999999;background-color: #D7FFEE">
					Date
				</td>
				<td width='80%' height='23' valign="middle" style="border: 1px solid #999999;">
					&nbsp;&nbsp;&nbsp;&nbsp;<%=docDate%></td>
			</tr>


		</table>

	</body>
</HTML>
