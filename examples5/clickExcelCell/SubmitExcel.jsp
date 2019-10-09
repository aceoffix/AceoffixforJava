<%@ page language="java"
	import="java.util.*,com.acesoft.aceoffix.*,com.acesoft.aceoffix.excelwriter.*"
	pageEncoding="UTF-8"%>
<%
	AceoffixCtrl aceCtrl = new AceoffixCtrl(request);
	aceCtrl.setServerPage("../server.ace");//Required

	Workbook workBook = new Workbook();
	
	Sheet sheet = workBook.openSheet("Sheet1");

	Table table = sheet.openTable("B4:D8");
	
	table.setReadOnly(false);
	
    aceCtrl.setJsFunction_OnExcelCellClick("OnCellClick()");
    
	aceCtrl.addCustomToolButton("Save", "Save", 1);

	aceCtrl.setSaveDataPage("SaveData.jsp");

	aceCtrl.bind(workBook);
	
	aceCtrl.setMenubar(false);
	
	aceCtrl.openDocument("doc/test.xls",OpenModeType.xlsSubmitForm,"Tom");
	 
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

<title>Acesoft Corporation - Capture the click event on the cells in Excel</title>

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
        <div class="left"><span>Current Page: <a href="../index.jsp" title="Examples Center">Example Center</a> &gt;</span><span><strong><em>Capture the click event on the cells in Excel</em></strong></span></div>
        <div style="text-align: right;">
            
        </div>
     
    </div>

<div class="container left">
	
<!--************** Start of AceoffixCtrl JavaScript ************************-->
<script type="text/javascript">
			function Save() {
				document.getElementById("AceoffixCtrl1").SaveDocument();
			}
			
			function OnCellClick(Celladdress, value, left, bottom) {
		            var i = 0;
		           
		            while (i<5) {
		      
		                if (Celladdress == "Sheet1!$B$" + (4 + i)) {
		                    var strRet = document.getElementById("AceoffixCtrl1").ShowHtmlModalDialog("select.jsp", "", "left=" + left + "px;top=" + bottom + "px;width=320px;height=180px;frame=no;");
		                    if (strRet != "") {
		                        return (strRet);
		                    }
		                    else {
		                        if ((value == undefined) || (value == ""))
		                            return " ";
		                        else
		                            return value;
		                    }
		                }
		                i++;
		            }
		        }
		</script>
<!--************** End of AceoffixCtrl JavaScript ************************-->	
		<form id="form1">
			Click the cells in Excel, a custom dialog box will prompt. You can only select "Department name" from a drop-down box.<br /><br />
			<div style="width: auto; height: 700px;">
				 <%= aceCtrl.getHtmlCode("AceoffixCtrl1")%>
			</div>
		</form>
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