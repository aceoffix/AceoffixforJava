<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="java.sql.*, javax.sql.*, java.io.*, sun.misc.*, java.math.*" %>
<%
String sID = request.getParameter("id");
if(sID != null){
	
	Class.forName("org.sqlite.JDBC"); 
	String strUrl ="jdbc:sqlite:" + this.getServletContext().getRealPath("data/")+ "/demo.db";
	Connection con = DriverManager.getConnection(strUrl);
	Statement stmt = con.createStatement();  
	String sql = "select * from Documents where ID = " + sID;		
	ResultSet rs = stmt.executeQuery(sql);
	
	if (rs.next()) {
	byte[] imageBytes = rs.getBytes("FileBin");
		int fileSize = imageBytes.length;

		response.reset();
		response.setContentType("application/msword");
		response.setHeader("Content-Disposition",
						"attachment; filename=stream.doc"); 
		response.setContentLength(fileSize);

		OutputStream outputStream = response.getOutputStream();
		outputStream.write(imageBytes);

		outputStream.flush();
		outputStream.close();
		outputStream = null;
		}
		stmt.close();
		con.close();
}
%>
