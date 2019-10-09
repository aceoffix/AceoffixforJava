<%@ page language="java"
	import="java.util.*,com.acesoft.aceoffix.*,java.sql.*"
	pageEncoding="UTF-8"%>
<%@page import="java.awt.image.ConvolveOp"%>
<%
	FileRequest freq = new FileRequest(request, response);
	String err = "";
	if (request.getParameter("id") != null
			&& request.getParameter("id").trim().length() > 0) {
		String id = request.getParameter("id").trim();
		Class.forName("org.sqlite.JDBC");
		String strUrl = "jdbc:sqlite:"
				+ this.getServletContext().getRealPath("data/")
				+ "/ExaminationPaper.db";
		Connection conn = DriverManager.getConnection(strUrl);
		String sql= "UPDATE  stream SET Word=?  where ID=" + id ;
		PreparedStatement pstmt=null;
		pstmt= conn.prepareStatement(sql);
		pstmt.setBytes(1,freq.getFileBytes());
		//pstmt.setBinaryStream(1,freq.getFileStream(),freq.getFileSize());
		pstmt.executeUpdate();
		pstmt.close();
		conn.close();

		freq.setCustomSaveResult("ok");
	} else {
		err = "<script>alert(' Save Failure');</script>";
	}
	freq.close();
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
		<%=err%>
	</body>
</html>
