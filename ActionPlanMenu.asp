<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(1) = false then
		Response.Redirect ("restricted.asp")
		Response.end
	end if

	PageTitle = "Online Management System!"
	PageName  = "ActionPlanMenu.asp"
%>
	
<!-- #Include file="include\header_menu.asp" -->

<%
	dim con, overwrite
	dim sqlDraft, sqlFinal, sqlExisting
	dim rsDraft, rsFinal, rsExisting

	set con	= server.createobject ("adodb.connection")		con.open "DSN=ehs"

	sqlDraft = "Select * from AP_ActionPlans where apFaculty = " & Session("DepID") & " and apCompleted = 0 order by apStartYear desc"
	set rsDraft = con.execute(sqlDraft)

	sqlFinal = "Select * from AP_ActionPlans where apFaculty = " & Session("DepID") & " and apCompleted = 1 order by apStartYear desc"
	set rsFinal = con.execute(sqlFinal)

	sqlExisting = "SELECT AP_ActionPlans.apStartYear " & _
				  "FROM AP_ActionPlans " & _
				  "WHERE (apStartYear = " & year(now()) & " or apCompleted = no) and apFaculty = " & Session("DepID") 
	set rsExisting = con.Execute(sqlExisting)

	' Determine if an Action Plan will be overwritten when a new Action Plan is created
	if not rsExisting.BOF then
		' Set the year of an existing Action Plan in the current year
		overwrite = rsExisting("apStartYear")
	else
		' There is no Action Plan to be overwritten
		overwrite = ""
	end if
%>

<table width="100%" cellspacing="0" border="0" cellpadding="4" align="center"><tr bgcolor="#0f4beb">
	<td>		<font size=+1 face=arial color=white>			<b>&nbsp; Health &amp; Safety Plans for the <% =Session("DepName") %></b>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<font size=-1 face=arial color=white><%=Session("AccessLevel")%></font>
        </font>	</td>
</tr></table><table align=center width="100%" cellspacing=4 border="0" cellpadding=4><tr>	<td valign=top width=35%>			<table bgcolor="#a0b8c8" width="100%" cellspacing="0" border="0" cellpadding="2">		<tr>			<td>
							<table bgcolor="white" width="100%" cellspacing="0" border="0" cellpadding="2">				<tr>					<td align="center">					
	                    <table width="100%" cellspacing="6" border="0" cellpadding="6">
						<!--<tr bgcolor="#eeeeee"> 
							<td valign="top"> 
								<b><font face="Arial" color="#000000">Getting Started</font></b><br>
								<b><font size="2" face="Arial" color="#3366cc">Faculties and high risk units:</font></b><br>
								<font size="2" face="Arial" color="#000000">Draft plan developed in H&amp;S Planning Workshop facilitated by Safety &amp; Wellbeing</font><br>
								<b><font size="2" face="Arial" color="#3366cc">Low risk units:</font></b><br>
								<font size="2" face="Arial">Identify high risk hazards/issues </font>
							</td>
						</tr>-->
<%
						if SecurityCheck(2) = true then
%>                       
						<!--<form name="ActionPlan" action="ActionPlanFormA.asp" method="post" target="_blank">
						<input type="hidden" name="exists" value="<%=overwrite%>">
						<input type="hidden" name="year" value="<%=year(now())%>">-->

						<tr bgcolor="#eeeeee"> 
							<td bgcolor="eeeeee">
								<b><font size="2" face="Arial" color="#3366cc">New Plan: </font></b><br>
								<font face="Arial" size="2" color="#000000">
									<!--<u style="cursor:pointer;cursor:hand;" onclick="javascript:NewPlan();">Create a new plan for <% =Session("DepName") %></u>-->
									<a href="javascript:void(0)" onclick="javascript:NewPlan('<%=SecurityCheck(4)%>');">Create a new plan for <% =Session("DepName") %></a>
								</font>
							</td>
						</tr>
						
<%
						End If
						
						if SecurityCheck(1) = true then
%>

						<!--</form>-->
						
						<tr bgcolor="#eeeeee"> 
							<td bgcolor="#eeeeee">
								<b><font size="2" face="Arial" color="#3366cc">Current Draft Plan:<br></font></b>
								<font color="#333333" size="2" face="Arial">
<%
								if not rsDraft.BOF then
									while not rsDraft.EOF
										if rsDraft("apFormACompleted") = true then
%>
										<!--<a href="ActionPlanFormB.asp?apID=<%=rsDraft("apID")%>" target="_blank"><%=rsDraft("apStartYear")%></a><BR>-->
										<!--<u style="cursor:pointer;cursor:hand;" onclick="javascript:OpenWindow('ActionPlanFormB.asp?apID=<%=rsDraft("apID")%>');"><%=rsDraft("apStartYear")%></u><BR>-->
										<a href="javascript:void(0)" onclick="javascript:OpenWindow('ActionPlanFormB.asp?apID=<%=rsDraft("apID")%>');"><%=rsDraft("apStartYear")%> - <%=rsDraft("apEndYear")%></a>&nbsp;<a href="#" onclick="checkDelete(<%=rsDraft("apID")%>)">Delete</a>
										
										<!-- displays a printer icon for printing a draft version of the ActionPlanFormB (EHS Plan) that shows all form field contents etc. CL 3/7/08 -->
										&nbsp;&nbsp;<a href="javascript:void(0)" onclick="javascript:OpenWindow('ActionPlanReportDraft.asp?apID=<%=rsDraft("apID")%>');"  title="Click on the printer icon to view a print-friendly version of the draft Environment, Health &amp; Safety Plan."><img src="printericon.gif" alt="Print-friendly format" width="16" height="16" border="0"></a>
										<!-- end of changes CL 3/7/08 -->
										
										<BR>
<%
										else 
%>
										<!--<a href="ActionPlanFormA.asp?apID=<%=rsDraft("apID")%>" target="_blank"><%=rsDraft("apStartYear")%></a><BR>-->
										<a href="javascript:void(0)" onclick="javascript:OpenWindow('ActionPlanFormA.asp?apID=<%=rsDraft("apID")%>');"><%=rsDraft("apStartYear")%> - <%=rsDraft("apEndYear")%></a>&nbsp;<a href="#" onclick="checkDelete(<%=rsDraft("apID")%>)">Delete</a><BR>
                                        
										<!--<u style="cursor:pointer;cursor:hand;" onclick="javascript:OpenWindow('ActionPlanFormA.asp?apID=<%=rsDraft("apID")%>');"><%=rsDraft("apStartYear")%></u><BR>-->
<%			
										end if 
										rsDraft.movenext
									wend
								end if
%>
								</font>                   
							</td>
						</tr>
<%
						end if
%>                        
						<tr bgcolor="#eeeeee"> 
							<td bgcolor="eeeeee"> 
								<b><font size="2" face="Arial" color="#3366cc">Final Plans:</font></b><br>
								<font color=333333 size=2 face=Arial>
<%
								if not rsFinal.BOF then
									while not rsFinal.EOF
%>
									<a href="javascript:void(0)" onclick="javascript:OpenWindow('ActionPlanReport.asp?apID=<%=rsFinal("apID")%>');"><%=rsFinal("apStartYear")%> - <%=rsFinal("apEndYear")%></a><BR>
<%
										rsFinal.movenext
									wend
								end if
%> 
								</font>
							</td>
						</tr>
						</table>
						
					</td>
				</tr>
				</table>
				
			</td>
		</tr>
		</table>

        <script type="text/javascript">
            function checkDelete(abc){
                var r = confirm("Are you sure you wish to delete this draft?");
                if (r == true) {
                    // Fire off the request to /form.php
                    request = $.ajax({
                        url: "AJAXDeleteDraft.asp",
                        type: "post",
                        data: "apid="+abc,
                        async: false,
                        success: function (data) {
                         
                            location.reload();
                        }
                    });
                } else {
                    //Don't need to do anything.
                }
            }
        </script>
		
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
							<td><font face="Arial, Helvetica, sans-serif" size="-1">Need more information on the UTS Health &amp; Safety Management System?<br><br>An outline of the system is available from the <a href="http://www.safetyandwellbeing.uts.edu.au/">Safety &amp; Wellbeing web site</a>.</font></td>
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
				<p><font face="Arial, Helvetica, sans-serif" size="-1"><b>Health &amp; Safety planning template</b><br>At any stage of the process users can:</font></p>
				
				<ul>
					<li><font face="Arial, Helvetica, sans-serif" size="-1">make amendments and save changes to the plan</font></li>
					<li><font face="Arial, Helvetica, sans-serif" size="-1">print the plan in hardcopy so that it can be circulated for consultation/discussion</font></li>
				</ul>
				
				<p><font face="Arial, Helvetica, sans-serif" size="-1">The Health &amp; Safety Plan is structured around a list of compliance requirements, which every faculty/unit must address to meet legislative obligations and policy-driven outcomes at UTS. For each compliance requirement:</font></p>
				
				<ul>
					<li><font face="Arial, Helvetica, sans-serif" size="-1">options for procedures and activities are provided as a means of complying.</font></li>
					<li><font face="Arial, Helvetica, sans-serif" size="-1">some procedures are flagged as mandatory. These are procedures that every faculty/unit must implement to meet legislative requirements.</font></li>
				</ul>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>

<script type="text/javascript">
<!--
	function NewPlan(blnEHSUser) {
		var Answer
		
		//if (blnEHSUser == "True"){
	
		if ("<%=overwrite%>" == "<%=year(now())%>") {
			
			if (blnEHSUser == "True") {
				Answer = confirm("There is already a draft or final Health and Safety Plan for this year.\n\nIf you proceed, the current HS Plan will be deleted as well as any associated Service Agreements, Facility Audits and Compliance Checks, and cannot be recovered.\n\nDo you want to proceed?");
				
				if(Answer) {
					OpenWindow("ActionPlanFormA.asp?exists=<%=overwrite%>&year=<%=year(now())%>")
				}
			}
			else {
				alert ("There is already an existing draft or final Plan for this year. If you wish to have the existing plan deleted please contact Safety &amp; Wellbeing.")
			}	
		}
		else {
		
			if ("<%=overwrite%>" != "") {
				if (blnEHSUser == "True") {
				
					Answer = confirm("There is already a draft or final EHS Plan for this year.\n\nI understand that if I proceed, the current EHS Plan will be deleted as well as any associated Service Agreements, Facility Audits and Compliance Assessments.  I have informed the relevant department that by deleting a EHS Plan that all related Service Agreements, Facility Audits and Compliance Assessments will be deleted and cannot be recovered.\n\nDo you want to proceed?");
				
					if(Answer) {
						OpenWindow("ActionPlanFormA.asp?exists=<%=overwrite%>&year=<%=year(now())%>")
					}
				}
				else {
					alert ("There is already an existing draft or final Plan for this year. If you wish to have the existing plan deleted please contact the Safety and Wellbeing Branch.")
				}
			}
			else {
				// There are is no draft and there is no final AP for the current year
			    OpenWindow("ActionPlanFormA.asp")
			}
		}		
	}
//-->
</script>

<!-- #Include file="include\footer.asp" -->