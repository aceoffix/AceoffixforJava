<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*,com.acesoft.aceoffix.excelreader.*, java.awt.*,java.text.*"%>

<%
	Workbook workBook = new Workbook(request, response);
	Sheet sheet = workBook.openSheet("Sheet1");
	
	String content = "";
	content += "testA:" + sheet.openCellByDefinedName("testA").getValue() + "<br/>";
    content += "testB:" + sheet.openCellByDefinedName("testB").getValue() + "<br/>";

	workBook.showPage(500, 400);
	workBook.close();
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
  <head>
    
    
    <title></title>
    
  </head>
  
  <body>
    <form id="form1">
			<div style="border: solid 1px gray;">
				<div class="errTopArea"
					style="text-align: left; border-bottom: solid 1px gray;">
					[Custom Dialog]
				</div>
				<div class="errTxtArea" style="height: 88%; text-align: left">
					<b class="txt_title">
						<div style=" color:#FF0000;" >
							Submit Information:
						</div> <%=content%> </b>
				</div>
				<div class="errBtmArea" style="text-align: center;">
					<input type="button" class="btnFn" value=" Close "
						onclick="window.opener=null;window.open('','_self');window.close();" />
				</div>
			</div>
		</form>
  </body>
</html>
