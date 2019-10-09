<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*, java.awt.*, com.acesoft.aceoffix.wordwriter.*"%>
<%
AceoffixCtrl aceCtrl = new AceoffixCtrl(request);
aceCtrl.setServerPage("../server.ace");//Required
aceCtrl.addCustomToolButton("Close","Close()",9);
aceCtrl.openDocument("doc/test.doc",OpenModeType.docNormalEdit,"Tom");
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html>
<head>
</head>
<body>
    <script type="text/javascript">
        function Close() {
           window.external.Close();
        }
        function increaseCount(value) {
		   var sResult = window.external.callParentJS("updateCount("+value+");");
		   if(sResult!="error:parentinvalid")//Note: The callParentJS() will return a invalid result when the parent page is refeshed or is closed or is redirected to other page.
			 document.getElementById("AceoffixCtrl1").Alert("The value of the Count in the parent page is: " + sResult);
        }
        function increaseCountAndClose(value) {
		   var sResult = window.external.callParentJS("updateCount("+value+");");
           window.external.Close();
        }
    </script>
 
    <input type="button" value="Add 1 to the value of the Count in the parent page" onclick="increaseCount(1);" />
	<input type="button" value="Add 5 to the value of the Count in the parent page, and then close the window." onclick="increaseCountAndClose(5);" /></br>
    <div style="width:auto; height:600px;">
      <%=aceCtrl.getHtmlCode("AceoffixCtrl1")%>
  </div>

</body>
</html>

