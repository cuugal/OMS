<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"            "http://www.w3.org/TR/html4/loose.dtd"><html><head><title>Online Management System - <% = PageTitle %></title><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><link rel="stylesheet" href="oms.css" type="text/css"><link rel="stylesheet" href="datatables/css/datatables.min.css" type="text/css"><link rel="stylesheet" href="datatables/css/jquery.ui.css" type="text/css"><link rel="stylesheet" href="datatables/TableTools/css/dataTables.tableTools.css" type="text/css"><script type="text/javascript" src="datatables/jquery.min.js"></script><script type="text/javascript" src="datatables/jquery-ui.js"></script><script type="text/javascript" src="datatables/datatables.min.js"></script><script type="text/javascript" src="dataTables/TableTools/js/dataTables.tableTools.min.js"></script><script type="text/javascript" src="datatables/moment.min.js"></script>
<script type="text/javascript" src="datatables/datetime-moment.js"></script><%	Session("MainURL") = PageName%>		<script type="text/javascript">		function OpenWindow(URL) 		{				window.open(URL)		}	</script></head><body bgcolor="#FFFFFF" text="#000000" style="width:970px; margin: 0 auto"><table border="0" cellpadding="2" width="100%" align="center"><tr height="150">	<td width="1%">		<table>					<tr>				<td><!-- commented out the old EHS Branch logo <img src="ehslogo2.gif" width="142" height="111" alt="EHS logo" border="0">--> &nbsp;</td>			</tr>		</table>	</td>	<td>			<table>					<tr height="50">				<td><br><br><br></td>			</tr>			<tr valign="middle" height="5">				<td nowrap valign="middle">					<div align="center">					<div align="center">						<font size="-1" face="arial">							<a href="Menu.asp">Home</a> &nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;							<a href="ActionPlanMenu.asp">Planning</a> 							<%if SecurityCheck(1) = true then%>							<!--a href="ServiceAgreementMenu.asp">Service Agreements</a--> <%							end if							%>								<%if SecurityCheck(1) = true then%>															&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp; <a href="AuditMenu.asp">Audits</a>							<%							end if							%>												</font>						<br>						</div>					<div align="center">					<font size="-1" face="arial">					<!-- a href="ComplianceMenu.asp">Compliance Assessments</a -->	<!-- show if super but not if auditor -->							<%if SecurityCheck(4) = true and SecurityCheck(5) = false then%>								<hr noshade size="0">								<a href="ManageAuditors.asp">Manage Auditors</a>							<%							end if							%><%					if SecurityCheck(4) = true  and SecurityCheck(5) = false then%>															 - <a href="ReportingMenu.asp">Management Reporting</a><%					end if%>							</font>						</div>			</div>			</td>		</tr>		</table>				</td>	<td width="1%">			<table>    <tr height="57"><td><br><br><br></td></tr>		<tr>    <td valign="middle" align="right">				<a href="http://www.uts.edu.au/"><img src="utslogo.gif" width="123" alt="The UTS home page" height="52" style="border:10px solid white" align="right"></a>			</td>      </tr>	  <tr><td>Logged in as: <% =Session("lgName") %> </td></tr><tr><td><a href="logout.asp" > Logout</a></td></tr>      </table>		</td>		</tr></table>