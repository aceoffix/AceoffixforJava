<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page language="java" import="java.io.*, java.sql.*, sun.misc.*" %>
<%@ page import="com.acesoft.aceoffix.*"%>
<%
String sID = request.getParameter("id");
if(sID != null){
	FileRequest freq = new FileRequest(request, response);
	byte[] b= freq.getFileBytes();
	Class.forName("org.sqlite.JDBC");
	String strUrl ="jdbc:sqlite:" + this.getServletContext().getRealPath("data/")+ "/demo.db";
	Connection con = DriverManager.getConnection(strUrl);
	Statement stmt = con.createStatement();
	PreparedStatement pstmt = null;
	pstmt = con.prepareStatement("update documents set FileBin=? where ID="+sID);
	pstmt.setBytes(1,b);
	pstmt.executeUpdate();
	pstmt.close();
	con.close();
	freq.close();
}
%>
