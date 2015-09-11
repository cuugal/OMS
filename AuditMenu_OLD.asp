<!-- #Include file="include\general.asp" -->

	if SecurityCheck(1) = false then
		Response.Redirect ("restricted.asp")
		Response.end
	end if
	
	PageTitle = "Health and Safety Online Management System"
	PageName  = "AuditMenu.asp"
	
<!-- #Include file="include\header_menu.asp" -->

	dim con, ActionPlan
	set con	= server.createobject ("adodb.connection")
		con.open "DSN=ehs"
		
	sqlAPID = "select apID " & _
			  "from AP_ActionPlans " & _
			  "where apFaculty = " & Session("DepID") & " and apStartYear = ( " & _
			  "		SELECT max(apStartYear) " & _
			  "		FROM AP_ActionPlans " & _
			  "		WHERE apFaculty = " & Session("DepID") & " and apCompleted = Yes )"
	set rsAPID = con.Execute(sqlAPID)
	if not rsAPID.BOF then

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

<table width=640 cellspacing="0" border="0" cellpadding="4" align="center">
	<td>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<font size="-1" face="arial" color="white"><%=Session("AccessLevel")%></font>
        </font>
</tr>
			
	                    <table width="100%" cellspacing="6" border="0" cellpadding="6">
	                    <tr bgcolor="#eeeeee">
							if ActionPlan <> "" then
%>
									<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditWorksheet.asp?apID=<%=ActionPlan%>');"><font color="#333333" size="2" face="Arial">Create an Audit Worksheet</font></a><br>
<%
							else
%>
									You need to complete a Plan before you can create a Facility Audit Worksheet
<%
							end if
%>                       
						</tr>
<%
						if SecurityCheck(2) = true then ' User must have write access for this department
%>
						<tr bgcolor="#eeeeee"> 
							<td bgcolor="#eeeeee">
							if ActionPlan <> "" then
%>
									<a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditForm.asp?apID=<%=ActionPlan%>&Mode=New');"><font color="#333333" size="2" face="Arial">Create a new Facility Audit Form</font></a><br>
<%
							else
%>
									<p>You need to complete a Plan before you can create a Facility Audit Form</p>
<%
							end if
%>
						</tr>
							<td bgcolor="#eeeeee">
								<b><font size="2" face="Arial" color="#3366cc">Current Draft Facility Audits:<br></font></b>
								<font color="#333333" size="2" face="Arial">
<%
								if not rsDraft.BOF then
									while not rsDraft.EOF
%>
<%
										rsDraft.movenext
									wend
								end if
%>
								</font>                   
							</td>
						</tr>
						<tr bgcolor="#eeeeee">
								<b><font size="2" face="Arial" color="#3366cc">Facility Audit Reports:<br></font></b>
								<font color="#333333" size="2" face="Arial">
							if not rsFinal.BOF then
								while not rsFinal.EOF
%>
<%
									rsFinal.movenext
								wend
							end if
%> 
							</td>
						</tr>
<%
						end if
%> 
	                    </table>
					</td>
				
		                
		                </tr>
		                <tr bgcolor="#eeeeee"> 
		                    <td><font face="Arial, Helvetica, sans-serif" size="-1">Need more information on the UTS Health &amp; Safety Management System?<br><br>An outline of the system is available from the <a href="http://www.safetyandwellbeing.uts.edu.au/">Safety &amp; Wellbeing Branch web site</a>.</font></td>
		                </tr>
		                </table>
					</td>
	
	
		<tr> 
			<td valign="top"> 
				<p><font face="Arial, Helvetica, sans-serif" size="-1"><b>Facility Audits </b><br><br>	
		</tr>
