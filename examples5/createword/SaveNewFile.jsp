<%@ page language="java"
	import="java.sql.*,java.io.*,javax.servlet.*,javax.servlet.http.*,java.text.SimpleDateFormat,java.util.Date,com.acesoft.aceoffix.*"
	pageEncoding="UTF-8"%>
<%
	FileRequest freq = new FileRequest(request, response);


	Class.forName("org.sqlite.JDBC");
	String strUrl = "jdbc:sqlite:"
			+ this.getServletContext().getRealPath("data/")
			+ "/CreateWord.db";
	Connection conn = DriverManager.getConnection(strUrl);
	Statement stmt = conn.createStatement();
	ResultSet rs = stmt.executeQuery("select Max(ID) from word");
	int newID = 1;
	if (rs.next()) {
		newID = Integer.parseInt(rs.getString(1)) + 1;
	}
	rs.close();
	
	String FileSubject = freq.getFormField("FileSubject").trim();
	String fileName = "test" + newID + ".doc"; 
	String getFile = (String) request.getParameter("FileSubject");
	if (getFile != null && getFile.length() > 0)
		FileSubject = new String(getFile.getBytes("UTF-8"));
	//out.print(FileSubject);
	SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	
	String strsql = "Insert into word(ID,FileName,Subject,SubmitTime) values("
			+ newID
			+ ",'"
			+ fileName
			+ "','"
			+ FileSubject
			+ "','"
			+ df.format(new Date()) + "')";
	stmt.executeUpdate(strsql);
	stmt.close();
	conn.close();
	
	freq.saveToFile(request.getSession().getServletContext().getRealPath("createword/doc/") + "/" + fileName);

	freq.setCustomSaveResult("ok");
	//fs.showPage(300,300);
	freq.close();
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
	<head>

		<title>My JSP 'SaveNewFile.jsp' starting page</title>

		<meta http-equiv="pragma" content="no-cache">
		<meta http-equiv="cache-control" content="no-cache">
		<meta http-equiv="expires" content="0">
		<meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
		<meta http-equiv="description" content="This is my page">

	</head>

	<body>
		<br>
	</body>
</html>
