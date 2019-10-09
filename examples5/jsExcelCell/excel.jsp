<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*,com.acesoft.aceoffix.excelwriter.*, java.awt.*,java.text.*"%>

<%
	
	AceoffixCtrl aceCtrl = new AceoffixCtrl(request);
	aceCtrl.setServerPage("../server.ace");//Required
	
	aceCtrl.setMenubar(false);
	
	aceCtrl.setCustomToolbar(false);
	
	aceCtrl.openDocument("doc/test.xls",OpenModeType.xlsNormalEdit,"Tom");
	 
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

<title>Acesoft Corporation - Use JavaScript to get and set the value in cells in Excel</title>

<!-- Theme styles -->
<link href="../css/acestyle.css" rel="stylesheet">
<link rel='stylesheet' id='contact-form-7-css'  href='../css/styles.css?ver=4.2.2' type='text/css' media='all' />
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
        <div class="left"><span>Current Page: <a href="../index.jsp" title="Examples Center">Example Center</a> &gt;</span><span><strong><em>Use JavaScript to get and set the value in cells in Excel</em></strong></span></div>
        <div style="text-align: right;">
            
        </div>
     
    </div>
<div class="container left">
	

<!--************** Start of AceoffixCtrl JavaScript ************************-->
<script type="text/javascript">
        function setCellValue(sheet, cell, value) {
            var sMac = "function myfunc()" + "\r\n"
                        + "Application.Sheets(\"" + sheet + "\").Range(\"" + cell + "\").Value = \"" + value + "\" \r\n"
                        + "End function";
            return document.getElementById("AceoffixCtrl1").RunMacro("myfunc", sMac);
        }
        function getCellValue(sheet, cell) {
            var sMac = "function myfunc()" + "\r\n"
                        + "myfunc = Application.Sheets(\"" + sheet + "\").Range(\"" + cell + "\").Text \r\n"
                        + "End function";
            return document.getElementById("AceoffixCtrl1").RunMacro("myfunc", sMac);
        }
        function Button1_onclick() {
            document.getElementById("AceoffixCtrl1").Alert(getCellValue("Sheet1", "D2"));
        }
        function Button2_onclick() {
            setCellValue("Sheet1", "F14", "100%");
        }
    </script>
<!--************** End of AceoffixCtrl JavaScript ************************-->
 Note: You need to set Macros in MS Excel first before clicking "Run Macro myfunc" button.<br/>
Steps for setting macros:<br/>
1.Open a Excel file manually.<br/>
2.Click "File"--"Options"--"Trust Center"--"Trust Center Settings"--"Macro Settings".<br/> 
3.Select "Trust access to the VBA project object model".<br/>
   Key code: right click and select "View Source" to check the function of JavaScript getCellValue(sheet, cell)&nbsp;&nbsp; setCellValue(sheet, cell, value).<br/>
	<div style="font-size:12px; line-height:20px; border-bottom:dotted 1px #ccc;border-top:dotted 1px #ccc; padding:5px;">
     <span style="color:red;">Operating:</span>click the button.
        <input id="Button1" type="button" value="Get the value of the cell D2 in Sheet1" onclick="return Button1_onclick()" />
        <input id="Button2" type="button" value="Set the value of F14 as 100% in Sheet1." onclick="return Button2_onclick()" />
     <br />
   
</div><br />

<div style="width:auto; height:800px;">
      <%= aceCtrl.getHtmlCode("AceoffixCtrl1")%>
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
