<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*,com.acesoft.aceoffix.wordreader.*,java.awt.*,javax.servlet.*,javax.servlet.http.*,java.sql.*,java.text.SimpleDateFormat,java.util.Date"%>

<%
	Class.forName("org.sqlite.JDBC");
	String strUrl = "jdbc:sqlite:" + this.getServletContext().getRealPath("data/") + "/demo2.db";
	Connection conn = DriverManager.getConnection(strUrl);
	Statement stmt = conn.createStatement();
	String id = request.getParameter("ID");
	String ErrorMsg = "";
	String BaseUrl = "";
	
	WordDocument doc = new WordDocument(request, response);
	String sName = doc.openDataRegion("name").getValue();
	String sDept = doc.openDataRegion("dept").getValue();
	String sCause = doc.openDataRegion("cause").getValue();
	String sDays = doc.openDataRegion("days").getValue();
	String sDate = doc.openDataRegion("date").getValue();

	if (sName.equals("")) {
		ErrorMsg = ErrorMsg + "<li>Name</li>";
	}
	if (sDept.equals("")) {
		ErrorMsg = ErrorMsg + "<li>Department</li>";
	}
	if (sCause.equals("")) {
		ErrorMsg = ErrorMsg + "<li>Cause</li>";
	}
	if (sDate.equals("")) {
		ErrorMsg = ErrorMsg + "<li>Date</li>";
	}
	try {
		if (sDays != "") {
			if (Integer.parseInt(sDays) < 0) {
				ErrorMsg = ErrorMsg + "<li>The number of the days can't be minus.</li>";
			}
		} else {
			ErrorMsg = ErrorMsg + "<li>Days</li>";
		}
	} catch (Exception Ex) {
		ErrorMsg = ErrorMsg	+ "<li><font color=red>Notice:</font>You should fill the 'Days' with numbers.</li>";
	}

	if (ErrorMsg == "") {
	//	String strsql = "update AbsenceRecord set Name='" + sName
	//			+ "', Dept='" + sDept + "', Cause='" + sCause
		//		+ "', Num=" + sDays + ", SubmitTime='" + sDate
			//	+ "' where  ID=" + id;
		String strsql = "update AbsenceRecord set Name='" + sName
				+ "', Dept='" + sDept + "', Cause='" + sCause
				+ "', Num=" + sDays + ", SubmitTime='" + sDate
				+ "' where  ID=" + id;
		//stmt.executeUpdate(strsql);	
		stmt.executeUpdate(strsql);
	} else {
		doc.showPage(578, 380);
	}
	doc.close();
	stmt.close();
	conn.close();
	
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<HTML>
	<HEAD>
		<title>SaveData</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">

	</HEAD>
	<body>
		<div class="errMainArea" id="error163"><div class="errTopArea" style="TEXT-ALIGN:left"></div>
			<div class="errTxtArea" style="HEIGHT:150px; TEXT-ALIGN:left">
				<b class="txt_title">
					<div style="color:#FF0000;">Please refill the "Days":</div>
					<ul>
					<%=ErrorMsg%>
					</ul>
					
				</b>
				
			</div>
			<div class="errBtmArea"><input type="button" class="btnFn" value="Close" onClick="window.opener=null;window.close();"></div>
		</div>
	</body>
</HTML>

