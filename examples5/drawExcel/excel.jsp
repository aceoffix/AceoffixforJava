<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*,com.acesoft.aceoffix.excelwriter.*, java.awt.*,java.text.*"%>

<%
	
	AceoffixCtrl aceCtrl = new AceoffixCtrl(request);
	aceCtrl.setServerPage("../server.ace");//Required
	
	Workbook wb = new Workbook();
	
	Table backGroundTable = wb.openSheet("Sheet1").openTable("A1:P200");
	backGroundTable.getBorder().setLineColor(Color.white);

	wb.openSheet("Sheet1").openTable("A1:H2").merge();
	wb.openSheet("Sheet1").openTable("A1:H2").setRowHeight(30);
	Cell A1 = wb.openSheet("Sheet1").openCell("A1");
	A1.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
	A1.setVerticalAlignment(XlVAlign.xlVAlignCenter);
	A1.setForeColor(new Color(255,64,0));
	A1.setValue("Consumption");
	
	wb.openSheet("Sheet1").openTable("A1:A1").getFont().setBold(true);
	wb.openSheet("Sheet1").openTable("A1:A1").getFont().setSize(25);
	
	//Border C4Border = wb.openSheet("Sheet1").openTable("C4:C4").getBorder();
	//C4Border.setWeight(XlBorderWeight.xlThick);
	//C4Border.setLineColor(Color.yellow);
	
	Table titleTable = wb.openSheet("Sheet1").openTable("B4:H5");
	titleTable.getBorder().setBorderType(XlBorderType.xlAllEdges);
	titleTable.getBorder().setWeight(XlBorderWeight.xlThick);
	titleTable.getBorder().setLineColor(new Color(255,64,0));
	
	Table bodyTable = wb.openSheet("Sheet1").openTable("B6:H15");
	bodyTable.getBorder().setLineColor(Color.gray);
	bodyTable.getBorder().setWeight(XlBorderWeight.xlHairline);

	Border B7Border = wb.openSheet("Sheet1").openTable("B7:B7").getBorder();
	B7Border.setLineColor(Color.white);

	Border B9Border = wb.openSheet("Sheet1").openTable("B9:B9").getBorder();
	B9Border.setBorderType(XlBorderType.xlBottomEdge);
	B9Border.setLineColor(Color.white);

	Border C6C15BorderLeft = wb.openSheet("Sheet1").openTable("C6:C15").getBorder();
	C6C15BorderLeft.setLineColor(Color.white);
	C6C15BorderLeft.setBorderType(XlBorderType.xlLeftEdge);
	
	Border C6C15BorderRight = wb.openSheet("Sheet1").openTable("C6:C15").getBorder();
	C6C15BorderRight.setLineColor(Color.yellow);
	C6C15BorderRight.setLineStyle(XlBorderLineStyle.xlDot);
	C6C15BorderRight.setBorderType(XlBorderType.xlRightEdge);

	Border E6E15Border = wb.openSheet("Sheet1").openTable("E6:E15").getBorder();
	E6E15Border.setLineStyle(XlBorderLineStyle.xlDot);
	E6E15Border.setBorderType(XlBorderType.xlAllEdges);
	E6E15Border.setLineColor(Color.yellow);

	Border G6G15BorderRight = wb.openSheet("Sheet1").openTable("G6:G15").getBorder();
	G6G15BorderRight.setBorderType(XlBorderType.xlRightEdge);
	G6G15BorderRight.setLineColor(Color.white);

	Border G6G15BorderLeft = wb.openSheet("Sheet1").openTable("G6:G15").getBorder();
	G6G15BorderLeft.setLineStyle(XlBorderLineStyle.xlDot);
	G6G15BorderLeft.setBorderType(XlBorderType.xlLeftEdge);
	G6G15BorderLeft.setLineColor(Color.yellow);

	Table bodyTable2 = wb.openSheet("Sheet1").openTable("B6:H15");
	bodyTable2.getBorder().setWeight(XlBorderWeight.xlThick);
	bodyTable2.getBorder().setLineColor(new Color(255,64,0));
	bodyTable2.getBorder().setBorderType(XlBorderType.xlAllEdges);

	Border H16H17Border = wb.openSheet("Sheet1").openTable("H16:H17").getBorder();
	H16H17Border.setLineColor(new Color(204, 255, 204));

	Border E16G17Border = wb.openSheet("Sheet1").openTable("E16:G17").getBorder();
	E16G17Border.setLineColor(new Color(255,64,0));

	Table footTable = wb.openSheet("Sheet1").openTable("B16:H17");
	footTable.getBorder().setWeight(XlBorderWeight.xlThick);
	footTable.getBorder().setLineColor(new Color(255,64,0));
	footTable.getBorder().setBorderType(XlBorderType.xlAllEdges);

	
	wb.openSheet("Sheet1").openTable("A1:A1").setColumnWidth(1);
	wb.openSheet("Sheet1").openTable("B1:B1").setColumnWidth(20);
	wb.openSheet("Sheet1").openTable("C1:C1").setColumnWidth(15);
	wb.openSheet("Sheet1").openTable("D1:D1").setColumnWidth(10);
	wb.openSheet("Sheet1").openTable("E1:E1").setColumnWidth(8);
	wb.openSheet("Sheet1").openTable("F1:F1").setColumnWidth(3);
	wb.openSheet("Sheet1").openTable("G1:G1").setColumnWidth(12);
	wb.openSheet("Sheet1").openTable("H1:H1").setColumnWidth(20);

	wb.openSheet("Sheet1").openTable("A16:A16").setRowHeight(20);
	wb.openSheet("Sheet1").openTable("A17:A17").setRowHeight(20);

	
	for (int i = 0; i < 12; i++) {
		for (int j = 0; j < 7; j++) {
			wb.openSheet("Sheet1").openCellRC(4 + i, 2 + j).getFont().setSize(10);
		}
	}

	
	for (int i = 0; i < 10; i++) {
		wb.openSheet("Sheet1").openCell("H" + (6 + i)).setBackColor(new Color(255, 255, 153));
	}

	wb.openSheet("Sheet1").openCell("E16").setBackColor(new Color(255,64,0));
	wb.openSheet("Sheet1").openCell("F16").setBackColor(new Color(255,64,0));
	wb.openSheet("Sheet1").openCell("G16").setBackColor(new Color(255,64,0));
	wb.openSheet("Sheet1").openCell("E17").setBackColor(new Color(255,64,0));
	wb.openSheet("Sheet1").openCell("F17").setBackColor(new Color(255,64,0));
	wb.openSheet("Sheet1").openCell("G17").setBackColor(new Color(255,64,0));
	wb.openSheet("Sheet1").openCell("H16").setBackColor(new Color(204, 255, 204));
	wb.openSheet("Sheet1").openCell("H17").setBackColor(new Color(204, 255, 204));

	
	Cell B4 = wb.openSheet("Sheet1").openCell("B4");
	B4.getFont().setBold(true);
	B4.setValue("Consumption");
	Cell H5 = wb.openSheet("Sheet1").openCell("H5");
	H5.getFont().setBold(true);
	H5.setValue("total");
	H5.setHorizontalAlignment(XlHAlign.xlHAlignCenter);
	Cell B6 = wb.openSheet("Sheet1").openCell("B6");
	B6.getFont().setBold(true);
	B6.setValue("airfare");
	Cell B9 = wb.openSheet("Sheet1").openCell("B9");
	B9.getFont().setBold(true);
	B9.setValue("hotel");
	Cell B11 = wb.openSheet("Sheet1").openCell("B11");
	B11.getFont().setBold(true);
	B11.setValue("repast");
	Cell B12 = wb.openSheet("Sheet1").openCell("B12");
	B12.getFont().setBold(true);
	B12.setValue("transportation fee");
	Cell B13 = wb.openSheet("Sheet1").openCell("B13");
	B13.getFont().setBold(true);
	B13.setValue("entertainment");
	Cell B14 = wb.openSheet("Sheet1").openCell("B14");
	B14.getFont().setBold(true);
	B14.setValue("gift");
	Cell B15 = wb.openSheet("Sheet1").openCell("B15");
	B15.getFont().setBold(true);
	B15.getFont().setSize(10);
	B15.setValue("other");

	wb.openSheet("Sheet1").openCell("C6").setValue("air fare");
	wb.openSheet("Sheet1").openCell("C7").setValue("air fare(back)");
	wb.openSheet("Sheet1").openCell("C8").setValue("other");
	wb.openSheet("Sheet1").openCell("C9").setValue("Daily National Consumption");
	wb.openSheet("Sheet1").openCell("C10").setValue("other");
	wb.openSheet("Sheet1").openCell("C11").setValue("Daily National Consumption");
	wb.openSheet("Sheet1").openCell("C12").setValue("Daily National Consumption");
	wb.openSheet("Sheet1").openCell("C13").setValue("total");
	wb.openSheet("Sheet1").openCell("C14").setValue("total");
	wb.openSheet("Sheet1").openCell("C15").setValue("total");

	wb.openSheet("Sheet1").openCell("G6");
	wb.openSheet("Sheet1").openCell("G7");
	wb.openSheet("Sheet1").openCell("G9");
	wb.openSheet("Sheet1").openCell("G10");
	wb.openSheet("Sheet1").openCell("G11").setValue(" day");
	wb.openSheet("Sheet1").openCell("G12").setValue(" day");

	wb.openSheet("Sheet1").openCell("H6").setFormula("=D6*F6");
	wb.openSheet("Sheet1").openCell("H7").setFormula("=D7*F7");
	wb.openSheet("Sheet1").openCell("H8").setFormula("=D8*F8");
	wb.openSheet("Sheet1").openCell("H9").setFormula("=D9*F9");
	wb.openSheet("Sheet1").openCell("H10").setFormula("=D10*F10");
	wb.openSheet("Sheet1").openCell("H11").setFormula("=D11*F11");
	wb.openSheet("Sheet1").openCell("H12").setFormula("=D12*F12");
	wb.openSheet("Sheet1").openCell("H13").setFormula("=D13*F13");
	wb.openSheet("Sheet1").openCell("H14").setFormula("=D14*F14");
	wb.openSheet("Sheet1").openCell("H15").setFormula("=D15*F15");

	for (int i = 0; i < 10; i++) {
		
		wb.openSheet("Sheet1").openCell("D" + (6 + i)).setNumberFormatLocal("$#,##0.00;$-#,##0.00");
		wb.openSheet("Sheet1").openCell("H" + (6 + i)).setNumberFormatLocal("$#,##0.00;$-#,##0.00");
	}

	Cell E16 = wb.openSheet("Sheet1").openCell("E16");
	E16.getFont().setBold(true);
	E16.getFont().setSize(11);
	E16.setForeColor(Color.white);
	E16.setValue("The total amount");
	E16.setVerticalAlignment(XlVAlign.xlVAlignCenter);
	Cell E17 = wb.openSheet("Sheet1").openCell("E17");
	E17.getFont().setBold(true);
	E17.getFont().setSize(11);
	E17.setForeColor(Color.white);
	E17.setFormula("=IF(C4>H16,\"budget\",\"budget\")");
	E17.setVerticalAlignment(XlVAlign.xlVAlignCenter);
	Cell H16 = wb.openSheet("Sheet1").openCell("H16");
	H16.setVerticalAlignment(XlVAlign.xlVAlignCenter);
	H16.setNumberFormatLocal("$#,##0.00;$-#,##0.00");
	H16.getFont().setName("Arial");
	H16.getFont().setSize(11);
	H16.getFont().setBold(true);
	H16.setFormula("=SUM(H6:H15)");
	Cell H17 = wb.openSheet("Sheet1").openCell("H17");
	H17.setVerticalAlignment(XlVAlign.xlVAlignCenter);
	H17.setNumberFormatLocal("$#,##0.00;$-#,##0.00");
	H17.getFont().setName("Arial");
	H17.getFont().setSize(11);
	H17.getFont().setBold(true);
	H17.setFormula("=(C4-H16)");

	
	Cell C4 = wb.openSheet("Sheet1").openCell("C4");
	C4.setNumberFormatLocal("$#,##0.00;$-#,##0.00");
	C4.setValue("2500");
	Cell D6 = wb.openSheet("Sheet1").openCell("D6");
	D6.setNumberFormatLocal("$#,##0.00;$-#,##0.00");
	D6.setValue("1200");
	wb.openSheet("Sheet1").openCell("F6").getFont().setSize(10);
	wb.openSheet("Sheet1").openCell("F6").setValue("1");
	Cell D7 = wb.openSheet("Sheet1").openCell("D7");
	D7.setNumberFormatLocal("$#,##0.00;$-#,##0.00");
	D7.setValue("875");
	wb.openSheet("Sheet1").openCell("F7").setValue("1");
	aceCtrl.bind(wb);
	
	aceCtrl.setMenubar(false);
	aceCtrl.setOfficeToolbars(false);
	
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

<title>Acesoft Corporation - Only using code to create a formatted Excel document from a blank file</title>

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
        <div class="left"><span>Current Page: <a href="../index.jsp" title="Examples Center">Example Center</a> &gt;</span><span><strong><em>Only using code to create a formatted Excel document from a blank file</em></strong></span></div>
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