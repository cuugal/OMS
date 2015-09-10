<!-- #Include file="include\general.asp" -->
<%
	if SecurityCheck(1) = false then
		Response.Redirect ("restricted.asp")
		Response.end
	end if
	
	PageTitle = "Health and Safety Online Management System"
	PageName  = "AuditMenu.asp"%>
	
<!-- #Include file="include\header_menu.asp" -->
<%
	dim con, ActionPlan	dim sqlDraft, sqlFinal, sqlAP	dim rsDraft, rsFinal, rsAP
	set con	= server.createobject ("adodb.connection")
		con.open "DSN=ehs"
			' Get the ID of the ActionPlan this Service Agreement will be attached to	
	sqlAPID = "select apID " & _
			  "from AP_ActionPlans " & _
			  "where apFaculty = " & Session("DepID") & " and apStartYear = ( " & _
			  "		SELECT max(apStartYear) " & _
			  "		FROM AP_ActionPlans " & _
			  "		WHERE apFaculty = " & Session("DepID") & " and apCompleted = Yes )"
	set rsAPID = con.Execute(sqlAPID)	
	if not rsAPID.BOF then		ActionPlan = rsAPID("apID")	else		ActionPlan = ""	end if

	sqlAP = "SELECT AP_ActionPlans.apStartYear, AP_ActionPlans.apID " & _
			"FROM AP_ActionPlans " & _
			"WHERE AP_ActionPlans.apFaculty = " & Session("DepID") & " " & _
			"ORDER BY AP_ActionPlans.apStartYear desc"
	set rsAP = con.Execute (sqlAP)

	sqlDraft = "SELECT apStartYear, apID, faComplete, faLabName, faID, faDate " & _
			   "FROM AP_ActionPlans INNER JOIN FA_Audits ON AP_ActionPlans.apID = FA_Audits.faActionPlan " & _
			   "WHERE  faComplete = No  AND apFaculty = " & Session("DepID") & " " & _
			   "ORDER BY AP_ActionPlans.apStartYear desc"
	set rsDraft = con.execute(sqlDraft)

	sqlFinal = "SELECT apStartYear, apID, faComplete, faLabName, faID, faDate " & _
			   "FROM AP_ActionPlans INNER JOIN FA_Audits ON AP_ActionPlans.apID = FA_Audits.faActionPlan " & _
			   "WHERE  faComplete = Yes  AND apFaculty = " & Session("DepID") & " " & _
			   "ORDER BY AP_ActionPlans.apStartYear desc"
	set rsFinal = con.execute(sqlFinal)
%>

<table width=640 cellspacing="0" border="0" cellpadding="4" align="center"><tr bgcolor="#6699cc">
	<td>		<font size="+1" face="arial" color="white">			<b>&nbsp; Facility Audits for the <% =Session("DepName") %></b>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<font size="-1" face="arial" color="white"><%=Session("AccessLevel")%></font>
        </font>	</td>
</tr></table><table align="center" width="640" cellspacing="4" border="0" cellpadding="4"><tr>	<td valign="top" width="35%">			<table bgcolor="#a0b8c8" width="100%" cellspacing="0" border="0" cellpadding="2">		<tr>			<td>
							<table bgcolor="white" width="100%" cellspacing="0" border="0" cellpadding="2">					<tr>					<td align="center">					
	                    <table width="100%" cellspacing="6" border="0" cellpadding="6">
	                    <tr bgcolor="#eeeeee">							<td bgcolor="#eeeeee">								<b><font size="2" face="Arial" color="#3366cc">Facility Audit Worksheet:</font></b><br><%
							if ActionPlan <> "" then
%>
									<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditWorksheet.asp?apID=<%=ActionPlan%>');"><font color="#333333" size="2" face="Arial">Create an Audit Worksheet</font></a><br>
<%
							else
%>
									You need to complete a Plan before you can create a Facility Audit Worksheet
<%
							end if
%>                       							</td>
						</tr>
<%
						if SecurityCheck(2) = true then ' User must have write access for this department
%>
						<tr bgcolor="#eeeeee"> 
							<td bgcolor="#eeeeee">								<b><font size="2" face="Arial" color="#3366cc">New Facility Audit Form:</font></b><br><%
							if ActionPlan <> "" then
%>
									<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditForm.asp?apID=<%=ActionPlan%>&Mode=New');"><font color="#333333" size="2" face="Arial">Create a new Facility Audit Form</font></a><br>
<%
							else
%>
									<p>You need to complete a Plan before you can create a Facility Audit Form</p>
<%
							end if
%>	                          </td>
						</tr>                    						<tr bgcolor="#eeeeee"> 
							<td bgcolor="#eeeeee">
								<b><font size="2" face="Arial" color="#3366cc">Current Draft Facility Audits:<br></font></b>
								<font color="#333333" size="2" face="Arial">
<%
								if not rsDraft.BOF then
									while not rsDraft.EOF
%><%=rsDraft("faDate")%>: <a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditForm.asp?apID=<%=rsDraft("apID")%>&Mode=Edit&faID=<%=rsDraft("faID")%>');">  <%=rsDraft("faLabName")%></a><!-- Makes a print-friendly version of the draft audit report available for the Admin user; click on the printer icon - CL 02/07/2008 -->&nbsp;&nbsp;<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditReportDraft.asp?apID=<%=rsDraft("apID")%>&faID=<%=rsDraft("faID")%>');" title="Click on the printer icon to view a print-friendly version of the draft audit report."><img src="printericon.gif" alt="Print-friendly format" width="16" height="16" border="0"></a> <!-- end of print-friendly changes --><BR>
<%
										rsDraft.movenext
									wend
								end if
%>
								</font>                   
							</td>
						</tr>
						<tr bgcolor="#eeeeee">							<td bgcolor="#eeeeee"> 
								<b><font size="2" face="Arial" color="#3366cc">Facility Audit Reports:<br></font></b>
								<font color="#333333" size="2" face="Arial"><%
							if not rsFinal.BOF then
								while not rsFinal.EOF
%><%=rsFinal("faDate")%>: <a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditReport.asp?apID=<%=rsFinal("apID")%>&faID=<%=rsFinal("faID")%>');"><%=rsFinal("faLabName")%></a><br>
<%
									rsFinal.movenext
								wend
							end if
%> 								</font>
							</td>
						</tr>
<%
						end if
%> 
	                    </table>	                    
					</td>				</tr>				</table>
							</td>		</tr>		</table>			<br>			<table bgcolor="#a0b8c8" width="100%" cellspacing="0" border="0" cellpadding="2">		<tr>			<td>							<table bgcolor="#eeeeee" width="100%" cellspacing="0" border="0" cellpadding="2">				<tr>					<td bgcolor="white" align="center">
		                		                <table width="100%" cellspacing="6" border="0" cellpadding="6">						<tr bgcolor="#eeeeee">  							<td valign="top">								<font face="Arial"><b>Need more information?</b></font>							</td>
		                </tr>
		                <tr bgcolor="#eeeeee"> 
		                    <td><font face="Arial, Helvetica, sans-serif" size="-1">Need more information on the UTS Health &amp; Safety Management System?<br><br>An outline of the system is available from the <a href="http://www.safetyandwellbeing.uts.edu.au/">Safety &amp; Wellbeing Branch web site</a>.</font></td>
		                </tr>
		                </table>
					</td>				</tr>				</table>			</td>		</tr>		</table>
		</td>	<td valign="top">
			<table width="100%" bgcolor="white" cellspacing="0" border="0"> 
		<tr> 
			<td valign="top"> 
				<p><font face="Arial, Helvetica, sans-serif" size="-1"><b>Facility Audits </b><br><br>					<U>Audit Worksheet</U><BR>				- a template to help conduct an audit on a facility or work area in your faculty/unit based on your Plan<br>				- used by the Safety &amp; Wellbeing Branch, but can be used by anyone at any point in time<br>				- can be printed off and used to note audit findings using audit criteria derived from your Plan<BR>				<BR>				<U>Audit Form</U><BR>				- online form used to enter audit results previously recorded on Audit Worksheet<BR>				- used by the Safety &amp; Wellbeing Branch, but can be used by Faculty/Unit management<BR>				- can be saved as 'draft' and returned to at any time<BR>				</font></td>
		</tr>		</table>	</td></tr></table>
<!-- #Include file="include\footer.asp" -->