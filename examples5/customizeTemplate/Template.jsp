<%@ page language="java" 
	import="java.util.*,com.acesoft.aceoffix.*,com.acesoft.aceoffix.wordwriter.*" 
	pageEncoding="UTF-8"%>

<%
		WordDocument doc = new WordDocument();
	
		doc.getTemplate().defineDataRegion("DATE", "[DATE]");
        doc.getTemplate().defineDataRegion("Payment", "[Payment method]");
        doc.getTemplate().defineDataRegion("Number", "[Credit card number]");
        doc.getTemplate().defineDataRegion("Status", "[Status]");
        doc.getTemplate().defineDataRegion("Name", "[Customer name]");
        doc.getTemplate().defineDataRegion("Add", "[Add]");
        doc.getTemplate().defineDataRegion("Tel", "[Tel]");
        doc.getTemplate().defineDataRegion("Email", "[Email]");
        	
        doc.getTemplate().defineDataTag("{Date}");
		
	
		AceoffixCtrl aceCtrl = new AceoffixCtrl(request);
		aceCtrl.setServerPage("../server.ace");//Required
        aceCtrl.addCustomToolButton("Save", "Save()", 1);
        aceCtrl.addCustomToolButton("Definition DataRegion", "ShowDefineDataRegions()", 3);
        aceCtrl.addCustomToolButton("Definition DataTags", "ShowDefineDataTags()", 3);
		aceCtrl1.addCustomToolButton("-", "", 0);
		aceCtrl1.addCustomToolButton("Close", "Close()", 9);

        
        aceCtrl.setSaveFilePage("SaveFile.jsp");

        aceCtrl.setTheme(ThemeType.Office2007);
        aceCtrl.setBorderStyle(BorderStyleType.BorderThin);
        aceCtrl.bind(doc);
        aceCtrl.openDocument("doc/test.doc", OpenModeType.docNormalEdit, "Tom");
        
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

<title>Acesoft Corporation - How user makes or customizes the Word template and how developer uses the template to generate the formatted Word document</title>

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
        <div class="left"><span>Current Page: <a href="../index.jsp" title="Examples Center">Example Center</a> &gt;</span><span><strong><em>How user makes or customizes the Word template and how developer uses the template to generate the formatted Word document</em></strong></span></div>
    </div>

<div class="container left">

<!--************** Start of AceoffixCtrl JavaScript ************************-->
	<script type="text/javascript">
        function getBkNames() {
            var bkNames = document.getElementById("AceoffixCtrl1").DataRegionList.DefineNames;
            return bkNames;
        }

        function getBkConts() {
            var bkConts = document.getElementById("AceoffixCtrl1").DataRegionList.DefineCaptions;
            return bkConts;
        }

        //Locate Bookmark
        function locateBK(bkName) {
            var drlist = document.getElementById("AceoffixCtrl1").DataRegionList;
            drlist.GetDataRegionByName(bkName).Locate();
            document.getElementById("AceoffixCtrl1").Activate();
            window.focus();


        }

        //Add Bookmark
        function addBookMark(param) {
            var tmpArr = param.split("=");
            var bkName = tmpArr[0];
            var content = tmpArr[1];
            var drlist = document.getElementById("AceoffixCtrl1").DataRegionList;
            drlist.Refresh();
            try {
                document.getElementById("AceoffixCtrl1").DocumentObject.Application.Selection.Collapse(0);
                drlist.Add(bkName, content);
                return "true";
            } catch (e) {
                return "false";
            }
        }
        //Delete Bookmark
        function delBookMark(bkName) {
            var drlist = document.getElementById("AceoffixCtrl1").DataRegionList;
            try {
                drlist.Delete(bkName);
                return "true";
            } catch (e) {
                return "false";
            }
        }

        function checkBkNames() {
            var drlist = document.getElementById("AceoffixCtrl1").DataRegionList;
            drlist.Refresh();
            var bkName = "";
            var bkNames = "";
            for (var i = 0; i < drlist.Count; i++) {
                bkName = drlist.Item(i).Name;
                bkNames += bkName.substr(4) + ",";
            }
            return bkNames.substr(0, bkNames.length - 1);
        }

        function checkBkConts() {
            var drlist = document.getElementById("AceoffixCtrl1").DataRegionList;
            drlist.Refresh();
            var bkCont = "";
            var bkConts = "";
            for (var i = 0; i < drlist.Count; i++) {
                bkCont = drlist.Item(i).Value;
                bkConts += bkCont + ",";
            }
            return bkConts.substr(0, bkConts.length - 1);
        }

        function getTagNames() {
            var tagNames = document.getElementById("AceoffixCtrl1").DefineTagNames;
            return tagNames;
        }

        //Locate Tag
        function locateTag(tagName) {

            var appSlt = document.getElementById("AceoffixCtrl1").DocumentObject.Application.Selection;
            var bFind = false;
            //appSlt.HomeKey(6);
            appSlt.Find.ClearFormatting();
            appSlt.Find.Replacement.ClearFormatting();

            bFind = appSlt.Find.Execute(tagName);
            if (!bFind) {
                document.getElementById("AceoffixCtrl1").Alert("The full text has been searched.");
                appSlt.HomeKey(6);
            }
            window.focus();

        }

        //Add Tag
        function addTag(tagName) {
            try {
                var tmpRange = document.getElementById("AceoffixCtrl1").DocumentObject.Application.Selection.Range;
                tmpRange.Text = tagName;
                tmpRange.Select();
                return "true";
            } catch (e) {
                return "false";
            }
        }

        //Delete Tag
        function delTag(tagName) {
            var tmpRange = document.getElementById("AceoffixCtrl1").DocumentObject.Application.Selection.Range;

            if (tagName == tmpRange.Text) {
                tmpRange.Text = "";
                return "true";
            }
            else
                return "false";
        }
   
        function Save() {
            document.getElementById("AceoffixCtrl1").SaveDocument();
        }
        function ShowDefineDataRegions() {
            document.getElementById("AceoffixCtrl1").ShowHtmlModelessDialog("DataRegionDlg.htm", "parameter=xx", "left=300px;top=390px;width=800px;height=410px;frame:no;");

        }
        function ShowDefineDataTags() {
            document.getElementById("AceoffixCtrl1").ShowHtmlModelessDialog("DataTagDlg.htm", "parameter=xx", "left=300px;top=390px;width=600px;height=410px;frame:no;");
        }
		function Close() {
			window.external.Close();
		}
    </script>
<!--************** End of AceoffixCtrl JavaScript ************************-->
    <form action="">
    <div style="width:auto; height:800px;">
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