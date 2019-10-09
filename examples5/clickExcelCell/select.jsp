<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%
String path = request.getContextPath();
String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/";
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
</head>
<body style="margin:0; padding:0;border-style:none; font-size:12px;overflow:hidden">
    <script language="javascript" type="text/javascript">
        function CheckValue(theForm) {
            window.external.DialogResult = theForm.WordList.value; 
            window.opener = null; 
            window.open('', '_self');
            window.close();
            return;
        }
	</script>
    <form id="form1" runat="server">
    <div  style=" margin:0; padding:0; height:10px; background-color:#D1E9E9; line-height:25px; text-align:right; padding:0 10px;">
        <a href="#" onclick="window.opener=null;window.open('','_self');window.close();" style="color:Black; font-size:18px; text-decoration:none;">×</a>
    </div>
    <div  id="rect1" style="background-color:#D1E9E9; margin:0; padding:0; height:150px;">
        <br />
       
        <br /><br />
<div style="text-align:center; ">
        <select name="WordList" style='width:240 px;'>
            <option value='Administration Department'>Administration Department</option>
            <option value='Accounting Department'>Accounting Department</option>
            <option value='Sales Department'>Sales Department</option>
            <option value='Marketing Department'>Marketing Department</option>
            <option value='Development Department'>Development Department</option>
            
        </select></div><br/>
        <div style="text-align:center; ">
        <button type="button" onclick="CheckValue(form1);">OK</button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <button type="button" onclick="window.opener=null;window.open('','_self');window.close();">Cancel</button>
        </div>
    </div>
    </form>
    <script language="javascript" type="text/javascript">
        document.getElementById("rect1").style.height = document.body.clientHeight - 30;
	</script>
</body>
</html>
