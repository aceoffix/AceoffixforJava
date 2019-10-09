<%@ page language="java" import="java.util.*,java.sql.*"
	pageEncoding="UTF-8"%>
<%
	Class.forName("org.sqlite.JDBC");
	String strUrl = "jdbc:sqlite:"
			+ this.getServletContext().getRealPath("data/")
			+ "/ExaminationPaper.db";
	Connection conn = DriverManager.getConnection(strUrl);
	Statement stmt = conn.createStatement();
	ResultSet rs = stmt.executeQuery("Select * from stream");
	boolean flg = false;
	StringBuilder strHtmls = new StringBuilder();
	while (rs.next()) {
		flg = true;
		String pID = rs.getString("ID");
		strHtmls.append("<tr  style='background-color:white;'>");
		strHtmls.append("<td><input name='check" + pID + "'  type='checkbox' /></td>");
		strHtmls.append("<td>Question-" + pID + "</td>");
		strHtmls.append("<td><a href=\"javascript:AceBrowser.openWindowModeless('Edit.jsp?id=" + pID + "','width=1200px;height=800px;');\">Edit</a></td>");
		strHtmls.append("</tr>");
	}

	if (!flg) {
		strHtmls.append("<tr>\r\n");
		strHtmls.append("<td width='100%' height='100' align='center'>Sorry \r\n");
		strHtmls.append("</td></tr>\r\n");
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

<title>Acesoft Corporation - Dynamically generate an examination paper in Word document</title>

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
        <div class="left"><span>Current Page: <a href="../index.jsp" title="Examples Center">Example Center</a> &gt;</span><span><strong><em>Dynamically generate an examination paper in Word document</em></strong></span></div>
        <div style="text-align: right;">
            
        </div>
     
    </div>

<div id="inner-wrap" class="clearfix">
	<div class="container">
    	<div class="inner-content">
	

<!--************** Start of AceoffixCtrl JavaScript ************************-->

    <script type="text/javascript">
		
		function button1Click(){
			document.getElementById("form1").action = "Compose.jsp";
			document.getElementById("form1").submit();
		}
		
		function button2Click(){
			document.getElementById("form1").action = "Compose2.jsp";
			document.getElementById("form1").submit();
		}
	</script>
<!--************** End of AceoffixCtrl JavaScript ************************-->	
		<div style="color: Red;text-align:center;">
			1.	Click "Edit" to edit each question.&nbsp;&nbsp;&nbsp;&nbsp;
		</div>
		<div style="color: Red;text-align:center;">
			2.	Select serveral questions and click "Generate Examination Paper" to generate the paper.
		</div>
		<form id="form1" name="form1" method="post" action="Compose.jsp" style="text-align: center;">
			<table style="background-color: #999; width: 600px;border: 1px solid #999999;" align="center">
				<%=strHtmls %>
			</table>
			<br />
			<div style="text-align:center;">
			<input type="button" onclick="button1Click();" id="Button1" value="Generate Examination Paper by JavaScript" /><span></span>
			<input type="button" onclick="button2Click();" id="Button2" value="Generate Paper by Java code" /><span style="color:Red;"></span></div>
		</form>
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
	
