<%@ page language="java"
	import="java.util.*, com.acesoft.aceoffix.*"
	pageEncoding="utf-8"%>
<%
AceoffixCtrl aceCtrl = new AceoffixCtrl(request);
aceCtrl.setServerPage("../server.ace");//Required
aceCtrl.addCustomToolButton("Save", "Save()", 1);
aceCtrl.addCustomToolButton("SaveAsPDF", "SaveAsPDF()", 1);
aceCtrl.setSaveFilePage("SaveFile.jsp");
String fileName = "editword.doc";
String pdfName = fileName.substring(0, fileName.length() - 4) + ".pdf";
aceCtrl.openDocument("doc/" + fileName, OpenModeType.docNormalEdit, "Tom");
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
	<head>
		<title>Convert Word to PDF</title>
		<script type="text/javascript">
        //保存
        function Save() {
            document.getElementById("AceoffixCtrl1").SaveDocument();
        }

        //另存为PDF文件
        function SaveAsPDF() {
            document.getElementById("AceoffixCtrl1").SaveDocumentAsPDF();
            document.getElementById("AceoffixCtrl1").Alert("The PDF file was saved in SaveAsPDF\\doc folder.");
        }
    </script>

	</head>
	<body>
		<form id="form1">
			<div id="div1"></div>
			<div style="width: auto; height: 700px;">
				 <%=aceCtrl.getHtmlCode("AceoffixCtrl1")%>
			</div>
		</form>
	</body>
</html>

