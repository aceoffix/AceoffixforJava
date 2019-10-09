<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*,com.acesoft.aceoffix.excelwriter.*, java.awt.*,java.text.*"%>

<%
	
	AceoffixCtrl aceCtrl = new AceoffixCtrl(request);
	aceCtrl.setServerPage("../server.ace");//Required
	
	Workbook wb = new Workbook();
	Sheet sheet = wb.openSheet("Sheet1");

	Cell cC3 = sheet.openCell("C3");
	
	cC3.setBackColor( Color.LIGHT_GRAY);
	cC3.setValue( "Jan");
	cC3.setForeColor(Color.white);
	cC3.setHorizontalAlignment(XlHAlign.xlHAlignCenter);

	Cell cD3 = sheet.openCell("D3");
	
	cD3.setBackColor( Color.lightGray);
	cD3.setValue( "Feb");
	cD3.setForeColor(Color.white);
	cD3.setHorizontalAlignment( XlHAlign.xlHAlignCenter);

	Cell cE3 = sheet.openCell("E3");
	
	cE3.setBackColor( Color.lightGray);
	cE3.setValue( "Mar");
	cE3.setForeColor(Color.white);
	cE3.setHorizontalAlignment( XlHAlign.xlHAlignCenter);

	Cell cB4 = sheet.openCell("B4");
	cB4.setBackColor(Color.BLUE);
	cB4.setValue( "Lodging");
	cB4.getFont().setSize(20);
	cB4.setForeColor(Color.YELLOW);
	cB4.setHorizontalAlignment( XlHAlign.xlHAlignCenter);

	Cell cB5 = sheet.openCell("B5");
	
	cB5.setBackColor( new Color(10,150,150));
	cB5.setValue( "Three Meals");
	cB5.setForeColor(Color.YELLOW);
	cB5.setHorizontalAlignment( XlHAlign.xlHAlignCenter);

	Cell cB6 = sheet.openCell("B6");
	
	cB6.setBackColor(new Color(200,200,100) );
	cB6.setValue( "Car Fare");
	cB6.setForeColor( new Color(10,150,150));
	cB6.setHorizontalAlignment( XlHAlign.xlHAlignCenter);

	Cell cB7 = sheet.openCell("B7");
	
	cB7.setBackColor( new Color(80,50,80));
	cB7.setValue( "Communication");
	cB7.setForeColor( new Color(255,90,10));
	cB7.setHorizontalAlignment( XlHAlign.xlHAlignCenter);

	
	Table titleTable = sheet.openTable("B3:E10");
	titleTable.getBorder().setWeight(XlBorderWeight.xlThick);
	titleTable.getBorder().setLineColor(new Color(255,64,0));
	titleTable.getBorder().setBorderType(XlBorderType.xlAllEdges);

	sheet.openTable("B1:E2").merge();
	sheet.openTable("B1:E2").setRowHeight( 30);
	Cell B1 = sheet.openCell("B1");
	
	B1.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
	B1.setVerticalAlignment(XlVAlign.xlVAlignCenter);
	B1.setForeColor( new Color(255,64,0));
	B1.setValue( "Consumption");
	B1.getFont().setBold(true);
	B1.getFont().setSize(25);
	

	aceCtrl.bind(wb);
	
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

<title>Acesoft Corporation - Set the font, color, alignment and background color of the text in the cells of Excel by programming</title>

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
        <div class="left"><span>Current Page: <a href="../index.jsp" title="Examples Center">Example Center</a> &gt;</span><span><strong><em>Set the font, color, alignment and background color of the text in the cells of Excel by programming</em></strong></span></div>
        <div style="text-align: right;">
            
        </div>
     
    </div>

<div class="container left">
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