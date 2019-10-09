<%@ page language="java"
	import="java.util.*,com.acesoft.aceoffix.*,com.acesoft.aceoffix.wordwriter.*,java.sql.*,java.io.*,javax.servlet.*,javax.servlet.http.*"
	pageEncoding="UTF-8"%>
<%!

public void copyFile(String oldPath, String newPath){

		try {
			int bytesum = 0;
			int byteread = 0;
			File oldfile = new File(oldPath);
			if (oldfile.exists()) { 
				InputStream inStream = new FileInputStream(oldPath); 
				FileOutputStream fs = new FileOutputStream(newPath);
				byte[] buffer = new byte[1444];
				int length;
				while ((byteread = inStream.read(buffer)) != -1) {
					bytesum += byteread;  
					
					fs.write(buffer, 0, byteread);
				}
				inStream.close();
			}
		} catch (Exception e) {
			System.out.println("error");
			e.printStackTrace();
		}

}
%>
<%

	String fileName = "";
	String mbName = request.getParameter("mb");

	AceoffixCtrl aceCtrl = new AceoffixCtrl(request);
	aceCtrl.setServerPage("../server.ace");//Required
	aceCtrl.setCustomToolbar(false);
	aceCtrl.setSaveFilePage("savefile.jsp");
	
	if (mbName != null && mbName.trim() != "") {
		
		
		
		fileName = "Official.doc";
		String templateName = request.getParameter("mb");
		String templatePath = getServletContext().getRealPath("applyTemplate/doc/" + templateName);
		String filePath = getServletContext().getRealPath("applyTemplate/doc/" + fileName);
		copyFile(templatePath, filePath); 

		
		WordDocument doc = new WordDocument();
		DataRegion  title= doc.openDataRegion("Title");
		title.setValue("Invitation Letter");
		DataRegion information = doc.openDataRegion("Information");
		information.setValue("Cooperative Extension is a publicly funded, non-formal educational partnership of the U.S. Department of Agriculture, land-grant universities, and county office.The USDA is an equal opportunity provider and employer.");
		DataRegion text = doc.openDataRegion("Text");
		text.setValue("[word]doc/test.doc[/word]");

		aceCtrl.bind(doc);
		
	} else {
		fileName = "test.doc";
		
	}
	
	aceCtrl.openDocument("doc/" + fileName, OpenModeType.docNormalEdit, "Tom");
	
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

<title>Acesoft Corporation - Apply the predefined template to current document by one click</title>

<!-- Theme styles -->
<link href="../css/acestyle.css" rel="stylesheet">
<link rel='stylesheet' id='contact-form-7-css'  href='../css/styles.css?ver=4.2.2' type='text/css' media='all' />
<!--************** Start of AceoffixCtrl JavaScript ************************-->
<script type="text/javascript">
	
	function load() {
		if (getQueryString("mb") != null)
			document.getElementById("templateName").value = getQueryString("mb");
	}

	function getQueryString(name) {
		var reg = new RegExp("(^|&)" + name + "=([^&]*)(&|$)", "i");
		var r = window.location.search.substr(1).match(reg);
		if (r != null)
			return unescape(r[2]);
		else
			return null;
	}

	function apply() {
		var mb = document.getElementById("templateName").value;
		document.getElementById("form1").action = "apply.jsp?mb=" + mb;

		document.forms[0].submit();
	}

	function saveAndClose() {
		document.getElementById("AceoffixCtrl1").SaveDocument();
		location.href = "index.jsp";
	}
</script>
<!--************** End of AceoffixCtrl JavaScript ************************-->	

</head>

<body onload="load();">
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
        <div class="left"><span>Current Page: <a href="../index.jsp" title="Examples Center">Example Center</a> &gt;</span><span><strong><em>Apply the predefined template to current document by one click</em></strong></span></div>
        <div style="text-align: right;">
            
        </div>
     
    </div>

<div class="container left">

	<div id="content">
			<div id="textcontent" style="width: 1000px; height: 800px;">
				<div class="flow4">
					<form method="post" id="form1">
						<strong>Template List:</strong>
						<span style="color: Red;"> <select name="templateName"
								id="templateName" style='width: 240px;'>
								<option value='temp2008.doc' selected="selected">
									Template I        
								</option>
								<option value='temp2009.doc'>
									Template II        
								</option>
								
							</select> </span>
						<span style="color: Red;"><input type="button" value="Use the Template"
								onclick="apply()"/> </span>
						<span style="color: Red;"><input type="button" value="Save and Close"
								onclick="saveAndClose()"/> </span>
					</form>
				</div>
				 <%= aceCtrl.getHtmlCode("AceoffixCtrl1")%>
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

