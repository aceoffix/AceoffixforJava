<%@ page language="java" 
	import="java.util.*, java.sql.*,java.io.*,javax.servlet.*,javax.servlet.http.*,java.text.SimpleDateFormat,java.util.Date" 
	pageEncoding="UTF-8"%>
<%
if(request.getParameter("op")!=null&&request.getParameter("op").length()>0){
          Class.forName("org.sqlite.JDBC");
	String strUrl = "jdbc:sqlite:"
		+ this.getServletContext().getRealPath("data/") + "/CreateWord.db";
	Connection conn = DriverManager.getConnection(strUrl);
	Statement stmt = conn.createStatement();
	ResultSet rs=stmt.executeQuery("select Max(ID) from word");
          int newID = 1;
          if (rs.next())
          {
              newID =  Integer.parseInt(rs.getString(1))+1;
          }
          rs.close();

          String fileName = "test" + newID + ".doc";
          String FileSubject = "Enter a Subject";
          String getFile=(String)request.getParameter("FileSubject");
          if ( getFile!=null&&getFile.length()>0)
              FileSubject =new String(getFile.getBytes("iso-8859-1"));
              out.print(FileSubject);
          SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	
          String strsql = "Insert into word(ID,FileName,Subject,SubmitTime) values(" + newID
              + ",'" + fileName + "','" + FileSubject + "','" + df.format(new Date()) + "')";
          stmt.executeUpdate(strsql);
          stmt.close();
          conn.close();
          
	String oldPath=getServletContext().getRealPath("createword/doc/template.doc");
	String newPath=getServletContext().getRealPath("createword/doc/" + fileName);
	try { 
       		int bytesum = 0; 
       		int byteread = 0; 
       		File oldfile = new File(oldPath);
       		if (oldfile.exists()) { 
           		InputStream inStream = new FileInputStream(oldPath); 
           		FileOutputStream fs = new FileOutputStream(newPath); 
           		byte[] buffer = new byte[1444]; 
           		int length; 
           		while ( (byteread = inStream.read(buffer)) != -1) { 
               		bytesum += byteread; 
               		fs.write(buffer, 0, byteread); 
           		} 
           		inStream.close(); 
       		} 
   		} 
  		catch (Exception e) { 
       		System.out.println("Error"); 
       		e.printStackTrace(); 
   		} 
		response.sendRedirect("word-lists.jsp");
	}
%>
  
    <!DOCTYPE HTML>
<!--[if lt IE 7]><html class="no-js lt-ie9 lt-ie8 lt-ie7"><![endif]-->
<!--[if IE 7]><html class="no-js lt-ie9 lt-ie8"><![endif]-->
<!--[if IE 8]><html class="no-js lt-ie9"><![endif]-->
<!--[if gt IE 8]><!--><html class="no-js"><!--<![endif]--><head>

<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>

<!-- Mobile Specific Metas -->
<meta name="viewport" content="width=device-width, initial-scale = 1.0, user-scalable = no">
<meta http-equiv="X-UA-Compatible" content="IE=Edge">

<title>Acesoft Corporation - Create document</title>

<!-- Theme styles -->
<link href="../css/acestyle.css" rel="stylesheet">
<link rel='stylesheet' id='contact-form-7-css'  href='../css/styles.css?ver=4.2.2' type='text/css' media='all' />
<!-- Aceoffix -->
<script type=text/javascript src="../jquery.min.js"></script>
<script type="text/javascript" src="../aceoffix.js" id="ace_js_main"></script>
</head>

<body>
<!--[if lt IE 7]>
	<p class="chromeframe">You are using an <strong>outdated</strong> browser. Please <a href="http://browsehappy.com/">upgrade your browser</a> or <a href="http://www.google.com/chromeframe/?redirect=true">activate Google Chrome Frame</a> to improve your experience.</p>
	<![endif]-->

<div id="banner-wrap" class="clearfix">
	<div id="top-nav" class="container clearfix">
        <div class="grid_10 center alpha">
        	Web Development Platform for Microsoft Office
        </div>
        <div class="grid_2 right omega">
        	<a href="http://www.aceoffix.com/index.php/contact-us/"><img src="../image/contactw.png" alt="contact">&nbsp;&nbsp;Contact Us</a>
        </div>
	</div>
    <div id="nav-wrap" class="clearfix">
        <div class="container">
            <div id="logo" class="grid_3 alpha"><a href="http://www.aceoffix.com"><img src="../image/aceoffixlogo.png" height="48" width="229" alt="Aceoffix" class="scale-with-grid" /></a></div>
            <div class="grid_9 omega">
                <nav class="nav">
                    <ul class="nav-list"><li class="page_item page-item-2 current_page_item menu-item-8"><a href="http://www.aceoffix.com/">Home</a></li>
<li class="page_item menu-item-244"><a href="http://www.aceoffix.com/index.php/overview/">Products</a>
</li>
<li class="page_item menu-item-224"><a href="http://www.aceoffix.com/index.php/download/">Download</a>
</li>
<li class="page_item menu-item-236"><a href="http://www.aceoffix.com/index.php/recent-update/">Support</a>
</li>
<li class="page_item menu-item-205"><a href="http://www.aceoffix.com/index.php/purchase/">Purchase</a>
</li>
<li class="page_item menu-item-17"><a href="http://www.aceoffix.com/index.php/about-us/">About Us</a></li>
</ul>                </nav>
            </div>
        </div>
    </div>


    <div id="title-wrap" class="clearfix">
        <div class="container left">
        Current Page:  Example Center
            <h1>Example Center - Aceoffix V5</h1>
            
        </div>
    </div>
    		
    </div>
    <br/>
    <div class="container left">
        <div class="left"><span>Current Page: <a href="../index.jsp" title="Examples Center">Example Center</a> &gt;</span><span><strong><em>Create document</em></strong></span></div>
        <div style="text-align: right;">
            
        </div>
     
    </div>

<div id="inner-wrap" class="clearfix">
	<div class="container">
    	<div class="inner-content">
<!--************** Start of AceoffixCtrl JavaScript ************************-->
<script type="text/javascript">
        function onColor(td) {
            td.style.backgroundColor = '#D7FFEE';
        }
        function offColor(td) {
            td.style.backgroundColor = '';
        }
        function getFocus() {
            var str = document.getElementById("FileSubject").value;
            if (str == "Enter a Subject") {
                document.getElementById("FileSubject").value = "";
            }
        }
        function lostFocus() {
            var str = document.getElementById("FileSubject").value;
            if (str.length <= 0) {
                document.getElementById("FileSubject").value = "Enter a Subject";
            }
        }
		function refreshPage() {
			setTimeout("window.location.reload()",1000);
		    return "ok";
		}
    </script>
<!--************** End of AceoffixCtrl JavaScript ************************-->
    <div class="zz-content mc clearfix pd-28">
    
        <div class="demo mc"  style=" text-align:center;">
            <h1 class="fs-16">
                There are two methods to create a document.</h1>
        </div>
         <div class="demo mc" style="text-align:center;">
            <form id="form1" method="post" action="word-lists.jsp?op=create">
            <center><table class="text" cellspacing="0" cellpadding="0" border="0">
                <tr>
                    <td style="font-size: 9pt" align="left">
                        1.	Create a file by copying a blank file.
                    </td>
                   <td align="center">
                        <input name="FileSubject" id="FileSubject" type="text" onfocus="getFocus()" onblur="lostFocus()"
                            class="boder" style="width: 180px;" value="Enter a Subject" />
                    </td>
                    <td style="width: 221px;">
                        &nbsp;
                        <input type="submit" value="Create" style=" width:86px;"/>
                    </td>
                </tr>
                <tr>
                    <td colspan="3">&nbsp;</td>
                </tr>
                <tr>
                    <td style="font-size: 9pt" align="left">
                        2.	Create a file by using CreateNewDocument method.
                    </td>
                    <td>
                        &nbsp;<a href="javascript:AceBrowser.openWindowModeless('CreateWord.jsp','width=1200px;height=800px;');" style=" color:Blue; text-decoration:underline;"> Create Document</a>
                        </td>
                    <td style="width: 221px;">
                        &nbsp;&nbsp;
                    </td>
                </tr>
            </table></center>
            </form>
        </div>
  
            <div class="zz-talbeBox mc" style=" text-align:center;" >
            <h1>Document List</h1>
           <center> <table class="zz-talbe" style="width: 80%;text-align: center;margin:auto">
                <thead>
                    <tr>
                        <th width="22%" style="border: 1px solid #999999;">
                            Subject    
                        </th>
                        <th width="10%" style="border: 1px solid #999999;">
                           Date
                        </th>
                    </tr>
                </thead>
                    <tbody style="border: 1px solid #999999;">
                    <%
                Class.forName("org.sqlite.JDBC");
				String strUrl = "jdbc:sqlite:"
					+ this.getServletContext().getRealPath("data/") + "/CreateWord.db";
				Connection conn = DriverManager.getConnection(strUrl);
				Statement stmt = conn.createStatement();
				ResultSet rs=stmt.executeQuery("select * from word order by id desc");
				String fileName="";
				String subject="";
				String submitTime="";
				while(rs.next()){	
					fileName = rs.getString("FileName");
					subject = rs.getString("Subject");
					submitTime = rs.getString("SubmitTime");
				%>
					<tr onmouseover='onColor(this)' onmouseout='offColor(this)'>
					<td style="border: 1px solid #999999;">
								<a href="javascript:AceBrowser.openWindowModeless('OpenWord.jsp?filename=<%= fileName %>&subject=<%=subject %>','width=1200px;height=800px;');"><%=subject %></a>
					</td>
				<%
					if(submitTime!=null&&submitTime.length()>0){
				%>
					<td style="border: 1px solid #999999;"><%=new SimpleDateFormat("MM/dd/yyyy")
								.format(new SimpleDateFormat("yyyy-MM-dd")
										.parse(submitTime)) %></td>
				<% 
					}else{
				%>
					<td>&nbsp;</td>
				<%
					}
				 %>
					</tr>
					<%
				}
				stmt.close();
				conn.close();
                     %>
                </tbody>
				
            </table></center>
        </div>
    </div>
 </div>
      
    </div>
</div><br/>
    
<div id="quicklink-wrap" class="clearfix">
	<div class="container">
    	<h1>Quick Link</h1>
       <div class="footer-column grid_3 alpha">
       		<h3>Product</h3>
			<a href="http://www.aceoffix.com/index.php/overview/">Overview</a>
<a href="http://www.aceoffix.com/index.php/functions-features/">Functions &#038; Features</a>
<a href="http://www.aceoffix.com/index.php/trial-license/">Trial &#038; License</a>
<a href="http://www.aceoffix.com/index.php/purchase/">Purchase</a>
       </div>
       <div class="footer-column grid_3 mobile-omega2">
			<h3>Download</h3>
			<a href="http://www.aceoffix.com/index.php/download/">Aceoffix for ASP.NET</a>
<a href="http://www.aceoffix.com/index.php/download/">Aceoffix for JAVA</a>
<a href="http://www.aceoffix.com/index.php/download/">Aceoffix for PHP</a>
       </div>
       <div class="footer-column grid_3 mobile-alpha2">
			<h3>Support</h3>
			<a href="http://www.aceoffix.com/index.php/technical-faq/">Technical FAQ</a>
<a href="http://www.aceoffix.com/index.php/videos/">Videos</a>
<a href="http://www.aceoffix.com/index.php/recent-update/">Recent Update</a>
       </div>
       <div class="footer-column grid_3 omega">
			<h3>Follow Us</h3>
			<a href="https://www.facebook.com/Aceoffix-152929314763836/">Facebook</a>
<a href="https://twitter.com/Aceoffix">Twitter</a>
       </div>
    </div>
</div>

<div id="footer-wrap" class="clearfix">
    <div class="container center">
		<a href="http://www.aceoffix.com/">Home</a>
<a href="http://www.aceoffix.com/index.php/acesoft-customer-privacy-policy/">Privacy Policy</a>
<a href="http://www.aceoffix.com/index.php/acesoft-terms-of-use/">Terms of Use</a>
<a href="http://www.aceoffix.com/index.php/contact-us/">Contact Us</a>
<a href="http://www.aceoffix.com/index.php/sitemap/">Sitemap</a>
		<br>
			&copy; 2018 Acesoft Corporation. All rights reserved.
    </div>
</div>

</body>
</html>      