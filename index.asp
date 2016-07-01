
<!--#include file="adovbs.inc"--> 
<%
	dim con 
	dim sqlPass
	dim rsPass 
	
	set rsPass = server.createobject("adodb.recordset")
	Set objCmd  = Server.CreateObject("ADODB.Command")		
	set con = server.CreateObject("adodb.connection")
	con.Open "DSN=ehs"
		
	Set objCmd.ActiveConnection = con	

	if Request("Login") <> "" then
		Session("Login") = Request("Login")
		Session("Pass") = Request("Password")
	
		' Compare the passwords and user name provided
		'AA 2014 - parametrise inputs to avoid injection 
		'sqlPass = "SELECT * FROM AD_Users WHERE lgName = '" & replace(Session("Login"),"'","''") & "' AND lgPassword = '" & replace(Session("Pass"),"'","''") & "'"
		sqlPass = "SELECT * FROM AD_Users WHERE lgName = ? AND lgPassword = ?"
			
		objCmd.CommandType = adCmdText
		objCmd.CommandText = sqlPass

		objCmd.Parameters.Append objCmd.CreateParameter("Login", adWChar, adParamInput, 50)
		objCmd.Parameters.Append objCmd.CreateParameter("Pass", adWChar, adParamInput,50)
		objCmd.Parameters("Login") = Session("Login")
		objCmd.Parameters("Pass") = Session("Pass") 
		
		'set rsPass = con.Execute (sqlPass)
		rsPass.Open objCmd
		
		if not rsPass.BOF then
			' Set the users Access Level
			Session("AccessLevel") = GetAccessLevel(rsPass("lgView"), rsPass("lgEdit"), rsPass("lgChangePassword"), rsPass("lgSuperUser"))
			Session("lgName") = rsPass("lgName")
			
			' Make sure here that the department is reset if we are re-logging in to the system for a different department
			Response.Redirect ("menu.asp?Department=" & rsPass("lgDepartment"))
			Response.End
		else
			Session("Login") = ""
			Session("Pass") = ""
%>
			<script type="text/javascript">
				alert("The login name or password is incorrect. Please check your login and try again.")
			</script>
<%
		end if

	end if
	
	function GetAccessLevel(View, Edit, Pass, Super)
		if Super = true then
			GetAccessLevel = "Super User"
		else
			if Pass = true then
				GetAccessLevel = "Manage"
			else
				if Edit = true then
					GetAccessLevel = "Edit User"
				else
					GetAccessLevel = "View User"
				End if
			End if
		end if
	end function
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
            "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<title>Online Management System (OMS)</title>
	<script type="text/javascript">
		// This is required to name the main window so it can be reloaded from child windows
		window.name='mainWindow'


  //This focuses the user client on the login ID form, for ease of login - CL 9/5/2011
  function formfocus() {
      document.getElementById('loginone').focus();
   }
   window.onload = formfocus;

 </script>
	<style type="text/css">
		.abbreviation table {font: Arial;}
		.abbreviation th { font-size: 90%; background-color: #0099CC; padding: 1px 4px; color: #fff; }
		.abbreviation td { font-size: 90%;  padding: 1px 14px; background-color: #eee; color: #000; text-align: left; }
		body { 	/* background: white url("ehslogo2.gif"); background-repeat: no-repeat; background-position: top right;*/ font: Arial, 10pt; padding: 20px; max-width: 85%; }
		table { font-family: Arial, sans-serif; font-size: 10pt; }
		p {text-align: justify; font-family: Arial, sans-serif; font-size: 10pt; margin-top: 0; padding-top: 0;}
		h4 { padding-bottom: 0; padding-top: 0; }
</style>
</head>

<body>
<center>
<table border="0" cellpadding="3">

<tr>
<!-- colour change from 0099CC for PROD to 669900 for DEV    Put in DEV in the Welcome title-->
	<td colspan="2" bgcolor="#0099CC" align="center"><h3><font color="white">Welcome to the Health and Safety Online Management System!</font></h3></td>
	<td><a href="http://www.uts.edu.au/"><img src="utslogo.gif" width="153" alt="The UTS home page" height="65" style="border:10px solid white" align="right"></a></td>
</tr>

<tr>
	<td width="15%" bgcolor="#eeeeee" valign="top" align="left">
		<b>Getting Started</b><br>
All UTS staff can access the OMS (view-only).<br><br>Simply enter the login ID and password for your area from the list at right.<br><br><strong>The password is the same as the login ID</strong>, and is not case-sensitive.<br><hr style="color: #a0b8c8; background-color: #a0b8c8; height: 1px; width: 25%; text-align: center; ">

	<form action="index.asp" name="login" method="post">    
		<font color="#333333" size="2"><b>Login ID</b></font><br>
		<input type="text" name="login" id="loginone" size="13"><br>         
		<font color="#333333" size="2"><b>Password</b></font><br>         
		<input type="password" name="password" size="13"><br><br>     
		<input type="submit" value="Sign in">
	</form>

<b>For More Information</b><br>
An outline of the University's Health and Safety Management System is available from the <a href="http://www.safetyandwellbeing.uts.edu.au/management/index.html" title="Outline of the UTS Health and Safety Management System at faculty/unit level">Safety &amp; Wellbeing web site</a>.</td>

	<td width="45%" valign="top" class="padding-left: 0;">
	<!-- start of logins -->
	<b>Login IDs for view-only access</b><br>
	<table class="abbreviation">
	<tbody>
	<tr>
		<th>Login ID</th>
		<th>Faculty/Unit/Institute</th>
	</tr>

	<tr>
		<td>BUS</td>
		<td>UTS Business School</td>
	</tr>

	<!--tr>
		<td>CIIC</td>
		<td>Creative Industries Innovation Centre</td>
	</tr-->

	<tr>
		<td>IPPG</td>
		<td>lnstitute for Public Policy and Governance</td>
	</tr>

	<tr>
		<td>DAB</td>
		<td>Faculty of Design, Architecture and Building</td>
	</tr>

<!--<tr>
		<td>EDU</td>
		<td>Faculty of Education [now FASS]</td>
	</tr>

	<tr>
		<td>ELSSA</td>
		<td>English Language and Study Skills Assistance Centre</td>
	</tr>
-->
	<tr>
		<td>EQDU</td>
		<td>Equity and Diversity Unit</td>
	</tr>

	<tr>
		<td>EXEC</td>
		<td>UTS Executive Group; including staff within the offices of the:
		<ul>
   <li>Vice-Chancellor</li>
   <li>Senior Deputy Vice-Chancellor</li>
   <li>Deputy Vice-Chancellor (Corporate Services)</li>
   <li>Deputy Vice-Chancellor (International & Development)</li>
   <li>Deputy Vice-Chancellor (Research)</li>
   <li>Deputy Vice-Chancellor (Resources)</li>
   <li>Deputy Vice-Chancellor (Teaching, Learning & Equity)</li>
   <li>Risk and Assurance Unit</li>
  </ul></td>
	</tr>

	<tr>
		<td>EXTREL</td>
		<td>External Relations Office, including: <ul><li>Alumni Relations</li><li>Development Office</li><li>Advancement Services</li><li>External Engagement</li></ul></td>
	</tr>
 
 <tr>
		<td>FASS</td>
		<td>Faculty of Arts and Social Sciences</td>
	</tr>

	<tr>
		<td>FEIT</td>
		<td>Faculty of Engineering and Information Technology - including:
		<ul><li>the Institute for Information and Communication Technologies (IICT)</li></ul></td>
	</tr>
<!--
	<tr>
		<td>FMU</td>
		<td>Facilities Management Unit, including:
		<ul><li>Commercial Services</li></ul></td>
	</tr>
-->
	<tr>
		<td>FMO</td>
		<td>Facilities Management Operations</td>
	</tr>
 
  <tr>
		<td>FSU</td>
		<td>Financial Services Unit</td>
  </tr>

	<tr>
		<td>GRS</td>
		<td>Graduate Research School</td>
	</tr>

  <tr>
		<td>GSH</td>
		<td>Graduate School of Health</td>
  </tr>
	
	<tr>
		<td>GSU</td>
		<td>Governance Support Unit</td>
	</tr>

	<tr>
		<td>HRU</td>
		<td>Human Resources Unit</td>
	</tr>

<!--	<tr>
		<td>HSS</td>
		<td>Faculty of Humanities and Social Sciences [now FASS]</td>
	</tr>
-->
	<tr>
		<td>ICI</td>
		<td>Innovation and Creative Intelligence</td>
	</tr>

	<tr>
		<td>IML</td>
		<td>Institute for Interactive Media and Learning</td>
	</tr>
<!--
	<tr>
		<td>IIS</td>
		<td>Institute for International Studies [now FASS]</td>
	</tr>
-->
	<tr>
		<td>IPPG</td>
		<td>lnstitute for Public Policy and Governance</td>
	</tr>

	<tr>
		<td>ISF</td>
		<td>Institute for Sustainable Futures</td>
	</tr>

<!--	<tr>
		<td>IT</td>
		<td>Faculty of Information Technology  [now FEIT] - including:
		<ul><li>the Institute for Information and Communication Technologies (IICT)</li></ul></td>
	</tr>
--> 
	<tr>
		<td>ITD</td>
		<td>Information Technology Division</td>
	</tr>

	<tr>
		<td>JIHL</td>
		<td>Jumbunna Indigenous House of Learning</td>
	</tr>

	<tr>
		<td>LAW</td>
		<td>Faculty of Law</td>
	</tr>

	<tr>
		<td>LIBRARY</td>
		<td>University Library</td>
	</tr>

<!--- <tr>
		<td>MCU</td>
		<td>Marketing and Communication Unit</td>
	</tr>
	--->

	<tr>
		<td>MCU</td>
		<td>Marketing and Communication Unit (including Events, Exhibitions and Projects)</td>
	</tr>

	<tr>
		<td>FOH</td>
		<td>Faculty of Health</td>
	</tr>

	<tr>
		<td>PMO</td>
		<td>Program Management Office</td>
	</tr>

	<tr>
		<td>PQU</td>
		<td>Planning and Quality Unit</td>
	</tr>

	<tr>
		<td>RIO</td>
		<td>Research and Innovation Office</td>
	</tr>

	<tr>
		<td>SAU</td>
		<td>Student Administration Unit</td>
	</tr>

	<tr>
		<td>SCI</td>
		<td>Faculty of Science - including:
		<ul><li>the Climate Change Cluster (C3)</li>
		<li>the &nbsp;Institute for Nanoscale Technology</li>
		<li>the &nbsp;Institute for the Biotechnology of Infectious Diseases (IBID)</li>
		</ul>
	</tr>

	<tr>
		<td>SHOPFRONT</td>
		<td>UTS Shopfront</td>
	</tr>

	<tr>
		<td>SSU</td>
		<td>Student Services Unit</td>
	</tr>

	<tr>
		<td>UTSC</td>
		<td>UTS Commercial</td>
	</tr>

	<tr>
		<td>UTSCC</td>
		<td>UTS Child Care</td>
	</tr>

	<tr>
		<td>LEGAL</td>
		<td>UTS Legal</td>
	</tr>

	<tr>
		<td>UTSI</td>
		<td>UTS: International</td>
	</tr>

	</tbody>
	</table>
	<!-- end of logins-->
	</td>

	<td width="45%" valign="top" align="left"><!-- start of overview -->
	<h4>Overview</h4>
	<p>The Online Management System (OMS) is an online application that supports the operational aspects of the University's health and safety management system. It does this by automatically generating templates for various stages of the planning cycle: PLANNING, SELF-ASSESSMENT and AUDIT.</p>

	<hr style="color: #a0b8c8; background-color: #a0b8c8; height: 1px; width: 25%; text-align: center; ">

	<h4>Health and Safety Planning</h4>
	<p>The University's health and safety management system requires all faculties and units to establish and maintain a current Health and Safety Plan. This is an operational plan that focuses on practical procedures to assist in compliance with our health and safety obligations.
	<br><br>
	An OMS electronic template is used to assist develop and draft your Plan at a workshop facilitated by the Safety &amp; Wellbeing Branch.
	<br><br>
	Each plan is created in such a way to allow scope for each faculty and unit to determine what procedures best suit their circumstances and the specific hazards relevant to their work environment as well as their courses and research, rather than imposing a "one-size fits all" manual of procedures.
	<br><br>
	The Safety &amp; Wellbeing Branch monitors and audits implementation of these plans and reports on implementation and compliance against the plans.</p>

	<!--hr style="color: #a0b8c8; background-color: #a0b8c8; height: 1px; width: 25%; text-align: center; ">

	<h4>Service Agreement</h4>
	<p>At the Planning Workshop, a Service Agreement is also negotiated with the Safety &amp; Wellbeing Branch. The aim of the Service Agreement is to provide each faculty or unit with the specific resources and support they consider useful to improve compliance. It outlines the services that the Safety &amp; Wellbeing Branch will provide to the faculty or unit to assist in the implementation of procedures developed in the plan. 
	<br><br>
	These services are recorded in an OMS electronic template for the purpose of tracking implementation progress.</p>

	<hr style="color: #a0b8c8; background-color: #a0b8c8; height: 1px; width: 25%; text-align: center; ">

	<h4>Self-assessment</h4>
	<p>At the midpoint of the plan duration, the Dean or Director undertakes a self-assessment of compliance of the whole faculty or unit against the Plan. This involves giving a numerical rating against each item in the Plan for that point in time.
	<br><br>
	These ratings are recorded in the OMS to act as a guide to how the faculty/unit estimates it is going and gives an opportunity to note any areas that require attention.</p-->

	<hr style="color: #a0b8c8; background-color: #a0b8c8; height: 1px; width: 25%; text-align: center; ">

	<h4>Audits</h4>
	<p>This section contains auditing templates with audit criteria derived from each faculty or unit Health and Safety Plan. The Safety &amp; Wellbeing Branch conducts regular audits of high-risk facilities within faculties and units using these audit criteria. Faculties and units can also use these to self-audit their own level of compliance against their Health and Safety Plan.<br><br>
	Audit results are recorded in the OMS and reports generated by the OMS are provided to the faculty/unit to facilitate continual improvement of health and safety management.</p>
	<!-- end of overview -->
	</td>

</tr>
</table>
<!-- end of new structure -->


</center>

</body>
</html>