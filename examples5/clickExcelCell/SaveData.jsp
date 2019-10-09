<%@ page language="java"
	import="java.util.*, java.text.*,com.acesoft.aceoffix.*,com.acesoft.aceoffix.excelreader.*"
	pageEncoding="UTF-8"%>
<%
	Workbook workBook = new Workbook(request, response);
	Sheet sheet = workBook.openSheet("Sheet1");
	Table table = sheet.openTable("B4:D8");
	String content = "";
	int result = 0;
	while (!table.getEOF()) {
	
        if (!table.getDataFields().getIsEmpty())
        {
            content += "<br/>Month:" + table.getDataFields().get(0).getText();
            content += "<br/>Plan:" + table.getDataFields().get(1).getText();
            content += "<br/>Reality:" + table.getDataFields().get(2).getText();

            content += "<br/>*********************************************";
        }
		
		table.nextRow();
	}
	table.close();
	

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
					
				</div>
				<div class="errTxtArea" style="height: 88%; text-align: left">
					<b class="txt_title">
						<div style=" color:#FF0000;" >
							Submit Information:
						</div> <%=content%> </b>
				</div>
				<div class="errBtmArea" style="text-align: center;">
					<input type="button" class="btnFn" value="Close"
						onclick="window.opener=null;window.open('','_self');window.close();" />
				</div>
			</div>
		</form>
	</body>
</html>