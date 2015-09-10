<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(1) = false then ' User 
		Response.Redirect ("restricted.asp")
		Response.end
	end if
	
	PageTitle = "EHS Online Management System!"
	PageName  = "ServiceAgreementMenu.asp"
%>
	
<!-- #Include file="include\header_menu.asp" -->

<%
	dim con, ActionPlan, SACount
	dim sqlDraft, sqlFinal
	dim rsDraft, rsFinal

	set con	= server.createobject ("adodb.connection")
		con.open "DSN=ehs"
	
	' Get the ID of the ActionPlan this Service Agreement will be attached to	
	sqlAPID = "select apID " & _
			  "from AP_ActionPlans " & _
			  "where apFaculty = " & Session("DepID") & " and apStartYear = ( " & _
			  "		SELECT max(apStartYear) " & _
			  "		FROM AP_ActionPlans " & _
			  "		WHERE apFaculty = " & Session("DepID") & " and apFormACompleted = Yes )"
	set rsAPID = con.Execute(sqlAPID)
	
	' Set the ActionPlan ID
	if not rsAPID.BOF then
		ActionPlan = rsAPID("apID")
		
		' Determine if there is a Draft ServiceAgreement or if a Service Agreement already exists for this ActionPlan
		sqlDraftCount = "SELECT count(*) as SACount " & _
						"FROM AP_ActionPlans INNER JOIN SA_ServiceAgreement ON AP_ActionPlans.apID = SA_ServiceAgreement.saActionPlan " & _
						"WHERE (saComplete = 0 AND apFaculty = " & Session("DepID") & ") or apID = " & ActionPlan
		set rsDraftCount = con.Execute(sqlDraftCount)
		
		saCount = rsDraftCount("SACount")
	else
		ActionPlan = null
		SACount = 0
	end if

	' Get the draft ServiceAgreement
	sqlDraft = "SELECT apID, apStartYear, saID " & _
			   "FROM AP_ActionPlans INNER JOIN SA_ServiceAgreement ON AP_ActionPlans.apID = SA_ServiceAgreement.saActionPlan " & _
			   "WHERE apFaculty = " & Session("DepID") & " AND saComplete = No"
	set rsDraft = con.execute(sqlDraft)

	' Get the final ServiceAgreements
	sqlFinal = "SELECT apID, apStartYear, saID " & _
			   "FROM AP_ActionPlans INNER JOIN SA_ServiceAgreement ON AP_ActionPlans.apID = SA_ServiceAgreement.saActionPlan " & _
			   "WHERE apFaculty = " & Session("DepID") & " AND saComplete = Yes"
	set rsFinal = con.execute(sqlFinal)
%>

<table width="100%" cellspacing="0" border="0" cellpadding="4" align="center">
<tr bgcolor="#6699cc">       
	<td>
		<font size="+1" face="arial" color="white"><b>&nbsp;Service Agreements for the <% =Session("DepName") %></b></font>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<font size="-1" face="arial" color="white"><%=Session("AccessLevel")%></font>
	</td>       
</tr>
</table>

<table width="100%" cellspacing="4" border="0" cellpadding="4" align="center">
<tr>
	<td valign="top" width="35%">
	
		<table bgcolor="#a0b8c8" width="100%" cellspacing="0" border="0" cellpadding="2">
		<tr>
			<td>
			
				<table bgcolor="white" width="100%" cellspacing="0" border="0" cellpadding="2">
				<tr>
					<td align="center">                     
					
						<table width="100%" cellspacing="6" border="0" cellpadding="6"> 
<% 	
						if SecurityCheck(4) = true then 
%>                                             
						<!--<form name="ServiceAgreement" action="ServiceAgreementForm.asp" method="get" target="_blank">
						<input type="hidden" name="exists" value="<% if SACount > 0 then Response.Write "yes" else Response.Write "no"%>">
						<input type="hidden" name="apID" value="<%=ActionPlan%>">-->

						<tr bgcolor="#eeeeee"> 
							<td bgcolor="#eeeeee">
								<b><font size="2" face="Arial" color="#3366cc">New Service Agreement: </font></b><br>
								<font face="Arial" size="2" color="#000000">
<%
									if isnull(ActionPlan) = false then
%>
									<a href="javascript:void(0)" onclick="javascript:NewPlan();">Create a new Service Agreement</u>
<%
									else
%>
									You need to Complete an Action Plan before you can create a Service Agreement
<%		
									end if
%>
								</font>
							</td>
						</tr>
						
						<!--</form>-->
<%
						end if
						
						if SecurityCheck(3) = true then 
%>						
						</form>
						
						<tr bgcolor="#eeeeee"> 
							<td bgcolor="#eeeeee">
								<b><font size="2" face="Arial" color="#3366cc">Current Draft Service Agreements:<br></font></b>
								<font color="#333333" size="2" face="Arial">
<%
								if not rsDraft.BOF then
%>
										<a href="javascript:void(0)" onclick="javascript:OpenWindow('ServiceAgreementForm.asp?apID=<%=rsDraft("apID")%>&saID=<%=rsDraft("saID")%>');"><%=rsDraft("apStartYear")%></a>
										
<!-- displays a printer-friendly icon for the printing the draft Service Agreement in admin mode only. CL 10/7/2008 -->
<a href="javascript:void(0)" onClick="javascript:OpenWindow('ServiceAgreementReportDraft.asp?apID=<%=rsDraft("apID")%>&saID=<%=rsDraft("saID")%>');" title="Click on the printer icon to view a print-friendly version of the Service Agreement."><img src="printericon.gif" alt="Print-friendly format" width="16" height="16" border="0"></a>
<!-- end of the printer-friendly Service Agreement section. CL 10/7/2008 -->

										<BR>
<%
								end if
%>
								</font>                   
							</td>
						</tr>
<%
						else
					
							if SecurityCheck(1) = true then 
%>						
						<tr bgcolor="#eeeeee"> 
							<td bgcolor="eeeeee">
								<b><font size="2" face="Arial" color="#3366cc">Current Draft Service Agreements:<br></font></b>
								<font color="#333333" size="2" face="Arial">
<%
								if not rsDraft.BOF then
%>
										<a href="javascript:void(0)" onclick="javascript:OpenWindow('ServiceAgreementReport.asp?apID=<%=rsDraft("apID")%>&saID=<%=rsDraft("saID")%>&draft=true');"><%=rsDraft("apStartYear")%></a><BR>
<%
								end if
%>
								</font>                   
							</td>
						</tr>
<%
							end if
						end if
						
						if SecurityCheck(4) = true then 
%>						
						<tr bgcolor="#eeeeee"> 
							<td bgcolor="eeeeee"> 
								<b><font size="2" face="Arial" color="#3366cc">Service Agreement Reports:</font></b><font face="Arial" size="2" color="#333333">
<%
								if not rsFinal.BOF then
									while not rsFinal.EOF
%>
									<br><a href="javascript:void(0)" onclick="javascript:OpenWindow('ServiceAgreementReport.asp?saID=<%=rsFinal("saID")%>&apID=<%=rsFinal("apID")%>');"><%=rsFinal("apStartYear")%></a>
<%
										rsFinal.movenext
									wend
								end if
%> 
								</font>
							</td>
						</tr>
<%                      else %>
							<tr bgcolor="#eeeeee"> 
							<td bgcolor="eeeeee"> 
								<b><font size="2" face="Arial" color="#3366cc">Service Agreement Reports:</font></b><font face="Arial" size="2" color="#333333">
<%
								if not rsFinal.BOF then
									while not rsFinal.EOF
%>
									<br><a href="javascript:void(0)" onclick="javascript:OpenWindow('ServiceAgreementReportNA.asp?saID=<%=rsFinal("saID")%>&apID=<%=rsFinal("apID")%>');"><%=rsFinal("apStartYear")%></a>
<%
										rsFinal.movenext
									wend
								end if
%> 
								</font>
							</td>
						</tr> <%
						end if
%>
						</table>
                  
					</td>
				</tr>
				</table>
				
			</td>
		</tr>
		</table>
		
		<br>
		
		<table bgcolor="#a0b8c8" width="100%" cellspacing="0" border="0" cellpadding="2">
		<tr>
			<td>
				
				<table bgcolor="#eeeeee" width="100%" cellspacing="0" border="0" cellpadding="2">
				<tr>
					<td bgcolor="white" align="center">                     
					
						<table width="100%" cellspacing="6" border="0" cellpadding="6">                       
						<tr bgcolor="#eeeeee">                          
							<td valign="top">
								<font face="Arial"><b>Need more information?</b></font>
							</td>
						</tr>
						<tr bgcolor="#eeeeee"> 
							<td><font face="Arial, Helvetica, sans-serif" size="-1">Need more information on the UTS Health and Safety Management System?<br><br>An outline of the system is available from the <a href="http://www.safetyandwellbeing.uts.edu.au/">Safety &amp; Wellbeing web site</a>.</font></td>
						</tr>
						</table>
						
					</td>
				</tr>
				</table>
				
			</td>
		</tr>
		</table>
	
	</td>
	<td valign="top">
	
		<table width="100%" bgcolor="white" cellspacing="0" border="0">
        <tr> 
			<td valign="top"> 
				<p><font face="Arial, Helvetica, sans-serif" size="-1"><b>Service Agreements</b></p>
				Service agreements are negotiated at Planning Workshops in consultation with key representatives from each faculty and unit. <br><br>The agreement records the specific services that Safety &amp; Wellbeing will provide to assist particular areas to meet their compliance requirements.
            </td>
        </tr>
        </table>
        
	</td>
</tr>
</table>

<script type="text/javascript">
<!--
	function NewPlan() {
		var Answer
	
		if ( <%=SACount%> > 0 ) {
			
			Answer = confirm("There is already a draft or final Service Agreement for the current EHS Plan. If you wish to proceed, this Service Agreement will be deleted and replaced by the new one. \n\nDo you want to proceed?");
			
			if(Answer) {
			     //document.ServiceAgreement.submit()
			     OpenWindow("ServiceAgreementForm.asp?exists=<% if SACount > 0 then Response.Write "yes" else Response.Write "no"%>&apID=<%=ActionPlan%>")
			}			
		}
		else {
			//document.ServiceAgreement.submit()
			OpenWindow("ServiceAgreementForm.asp?exists=<% if SACount > 0 then Response.Write "yes" else Response.Write "no"%>&apID=<%=ActionPlan%>")
		}
	}
//-->
</script>

<!-- #Include file="include\footer.asp" -->