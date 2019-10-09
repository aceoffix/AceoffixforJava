<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%@ page import="com.acesoft.aceoffix.*,com.acesoft.aceoffix.wordwriter.*,java.awt.*,javax.servlet.*,javax.servlet.http.*,java.sql.*,java.text.SimpleDateFormat,java.util.Date"%>
<%
	Class.forName("org.sqlite.JDBC");
	String strUrl = "jdbc:sqlite:" + this.getServletContext().getRealPath("data/")+ "/demo2.db";
	Connection conn = DriverManager.getConnection(strUrl);
	Statement stmt = conn.createStatement();
	String id = request.getParameter("ID");
	ResultSet rs = stmt.executeQuery("select * from AbsenceRecord where ID = " + id);

	String docSubject = "", docName = "", docDept = "", docCause = "", docDays = "", docDate = "", docFile = "";
	
	
	if (rs.next()) {
		docFile = rs.getString("FileName");
		docSubject = rs.getString("Subject");
		docName = rs.getString("Name");
		docDept = rs.getString("Dept");
		docCause = rs.getString("Cause");
		docDays = rs.getString("Num");
		docDate = rs.getString("SubmitTime");
	}
	rs.close();
	stmt.close();
	conn.close();

	WordDocument doc = new WordDocument();
	DataRegion drName = doc.openDataRegion("name");
	drName.setValue(docName);
	drName.setEditing(true);
	DataRegion drDept = doc.openDataRegion("dept");
	drDept.setValue(docDept);
	drDept.getShading().setBackgroundPatternColor(Color.GRAY);
	//drDept.setEditing(true);  
	DataRegion drCause = doc.openDataRegion("cause");
	drCause.setValue(docCause);
	drCause.setEditing(true);
	DataRegion drNum = doc.openDataRegion("days");
	drNum.setValue(docDays);
	drNum.setEditing(true);
	DataRegion drDate = doc.openDataRegion("date");
	drDate.setValue(docDate);
	//drDate.setValue(new SimpleDateFormat("MM/dd/yyyy").format(new SimpleDateFormat("yyyy-MM-dd").parse(docDate)));
	drDate.getShading().setBackgroundPatternColor(Color.pink);
	DataRegion toDate = doc.openDataRegion("ToDate");
	toDate.setEditing(true);
	//toDate.setValue(new SimpleDateFormat("MM/dd/yyyy").format(new SimpleDateFormat("yyyy-MM-dd").parse(docDate)));
	toDate.setValue(docDate);
	
	AceoffixCtrl aceCtrl = new AceoffixCtrl(request);
	aceCtrl.setServerPage("../server.ace");//Required

	aceCtrl.setBorderStyle(BorderStyleType.BorderThin);
	
	aceCtrl.addCustomToolButton("Save", "poSave()", 1);
	aceCtrl.addCustomToolButton("Full-screen Switch", "poSetFullScreen()", 4);
	
	aceCtrl.setOfficeToolbars(false);
	aceCtrl.setMenubar(false);

	aceCtrl.setJsFunction_OnWordDataRegionClick("OnWordDataRegionClick()");
	
	aceCtrl.setSaveDataPage("SaveData.jsp?ID=" + id);
	
	aceCtrl.bind(doc);
	
	aceCtrl.openDocument("doc/template.doc",OpenModeType.docSubmitForm,"Tom");
	 
	
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

<title>Acesoft Corporation - A demo of Absence Request</title>

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
        <div class="left"><span>Current Page: <a href="../index.jsp" title="Examples Center">Example Center</a> &gt;</span><span><strong><em>A demo of Absence Request</em></strong></span></div>
        <div style="text-align: right;">
            
        </div>
     
    </div>

<div class="container left">

<a href="index.jsp"> Back</a>
<!--************** Start of AceoffixCtrl JavaScript ************************-->
		<script language="JavaScript">
	function OnWordDataRegionClick(Name, Value, Left, Bottom) {
		if (Name == "ACE_date") {
	                var strRet = document.getElementById("AceoffixCtrl1").ShowHtmlModalDialog("datetimer.htm", Value, "left=" + Left + "px;top=" + Bottom + "px;width=320px;height=350px;frame=no;");
	                if (strRet != "") {
	                    return (strRet);
	                }
	                else {
	                    if ((Value == undefined) || (Value == ""))
	                        return " ";
	                    else
	                        return Value;
	                }
	            }
	            if (Name == "ACE_dept") {

	                var strRet = document.getElementById("AceoffixCtrl1").ShowHtmlModalDialog("selectDept.htm", Value, "left=" + Left + "px;top=" + Bottom + "px;width=270px;height=260px;frame=no;");
	                if (strRet != "") {
	                    return (strRet);
	                }
	                else {
	                    if ((Value == undefined) || (Value == ""))
	                        return " ";
	                    else
	                        return Value;
	                }
	            }
	}
</script>
<!--************** End of AceoffixCtrl JavaScript ************************-->
		<form id="Form1" method="post">

			<script type="text/javascript">
				
				function poSave() {
					 document.getElementById("AceoffixCtrl1").SaveDocument();
				}
				
				function poSetFullScreen() {
					document.getElementById("AceoffixCtrl1").FullScreen = !document.getElementById("AceoffixCtrl1").FullScreen;
				}
			</script>
			<div style="width:auto; height:600px;">
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
