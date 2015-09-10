<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(4) = false then 
		Response.Redirect ("restricted.asp")
		Response.end
	end if
	
	PageTitle = "EHS Online Management System!"
	PageName  = "ReportingMenu.asp"
%>
	
<!-- #Include file="include\header_menu.asp" -->

<%
	dim con, startYear, endYear

	set con	= server.createobject ("adodb.connection")
		con.open "DSN=ehs"

	dim sqlYear, sqlReq
	dim rsYear, rsReq
	
	' Determine the start and end years for this report
	sqlYear = "SELECT Min(apStartYear) AS minStart, Max(apStartYear) AS maxStart, datepart('yyyy', Min(ccDate)) AS minComp, datepart('yyyy', Max(ccDate)) AS maxComp " & _
			  "FROM CC_Compliance RIGHT JOIN AP_ActionPlans ON CC_Compliance.ccActionPlan = AP_ActionPlans.apID " & _
			  "WHERE apCompleted = Yes"
	set rsYear = con.Execute(sqlYear)

	'Response.Write isnull(rsYear("maxComp")) & "<BR>"
	'Response.Write rsYear("minComp") & "<BR>"
	'Response.Write rsYear("maxStart") & "<BR>"
	'Response.Write rsYear("maxComp") & "<BR>"
	
	
	if rsYear("minStart") < rsYear("minComp") or isnull(rsYear("minComp")) = true then
		startYear = rsYear("minStart")
	else
		startYear = rsYear("minComp")
	end if
	
	if rsYear("maxStart") < rsYear("maxComp") or isnull(rsYear("maxComp")) = true then
		endYear = rsYear("maxStart")
	else
		endYear = rsYear("maxComp")
	end if
	
	
	
	'Response.Write startyear & "<BR>"
	'Response.Write endYear & "<BR>"
	'Response.End
	
	sqlReq = "SELECT IN_Requirements.irId, IN_Requirements.irName " & _
			 "FROM IN_Requirements " & _
			 "WHERE irActive = Yes " & _
			 "ORDER BY IN_Requirements.irDisplayOrder"
	set rsReq = con.Execute(sqlReq)	
%>

<table width="100%" cellspacing="0" border="0" cellpadding="4" align="center">
<tr bgcolor="#6699cc">
	<td>
		<font size="+1" face="arial" color="white">&nbsp;<b>Management Reporting for the <% =Session("DepName") %></b></font>
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
						<tr bgcolor="#eeeeee">
							<td bgcolor="#eeeeee">
								<b><font face="Arial" color="#000000">Choose one of the reports below</font></b>
							</td>                       
						</tr>                       
						<tr bgcolor="#eeeeee">     
						
						    <form name="reqRpt" action="ReportingByRequirement.asp" method="post" target="_blank">               
							
							<td bgcolor="#eeeeee">
                <b><font size="2" face="Arial" color="#3366cc">Select a compliance requirement:</font></b>
                <br><br>
                <font face="Arial" size="2" color="#000000">
								
								<select name="req" onchange="javascript:this.form.submit();">
									<option value="-1">- Select a Requirement -</option>
<%
									while not rsReq.eof
%>
									<option value="<%=rsReq("irID")%>"><%=rsReq("irName")%></option>
<%
										rsReq.movenext
									wend
%>
								</select>

								</font>
							</td>
                        
							</form>
							
						</tr>
						<tr bgcolor="#eeeeee">
						
							<form name="yearRpt" action="ReportingByYear.asp" method="post">
						 
							<td bgcolor="#eeeeee">
								<b><font size="2" face="Arial" color="#3366cc">Select a year:</font></b><br><br><font face="Arial" size="2" color="#000000">

								<select name="year" onchange="javascript:this.form.submit();">
									<option value="-1">- Select a Year -</option>
<%
									dim curYear
	
									curYear = startYear

									while curYear <= endYear
%>
									<option value="<%=curYear%>"><%=curYear%></option>
<%
										curYear = curYear + 1
									wend
%>
								</select>
    
								</font>
        <br>
        <small><b>NOTE:</b> The compliance ratings shown in the "Select a Year" report are only derived from Planning Workshops (and <u>NOT</u> also from Compliance Assessments, as was previously the case).</small>
							</td>
                          
							</form>
                          
						</tr>
						<tr bgcolor="#eeeeee">
							<td bgcolor="#eeeeee">
								<font face="Arial" size="2" color="#3366CC"><b>Service Agreement Report</b><br>
								<br>Select month &amp; year from:</font>
								
								<form name="smgmt" action="SAMgmtReport.asp" method="POST" target ="_blank" >

									<p><font face="Arial" size="2">Month&nbsp;&nbsp;
									</font>&nbsp;<select size="1" name="cboMonth">
									<option value="1">January</option>
									<option value="2">February</option>
									<option value="3">March</option>
									<option value="4">April</option>
									<option value="5">May</option>
									<option value="6">June</option>
									<option value="7">July</option>
									<option value="8">August</option>
									<option value="9">September</option>
									<option value="10">October</option>
									<option value="11">November</option>
									<option value="12">December</option>
									</select><br>
									<br>
									<font face="Arial" size="2">Year </font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<!"Validation" s-data-type="String" b-value-required="TRUE" i-minimum-length="4" i-maximum-length="4" -->
									<input type="text" name="txtYear" size="4" style="font-family: Arial; font-size: 1em" maxlength="4"><font face="Arial" size="1">&nbsp;&nbsp; 
									(yyyy&nbsp; eg. 2008)<br>
									<br>
									</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<input type="button" value="Show Report" name="B1" onclick="javascript:this.form.submit();"></p>
								</form>
							</td>
                          
						</tr>
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
							<td><font face="Arial, Helvetica, sans-serif" size="-1">Need more information on the UTS Health and Safety Management System?<br><br>An outline of the system is available from <a href="http://www.safetyandwellbeing.uts.edu.au/">Safety &amp; Wellbeing</a>.</font></td>
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
			<td valign="top"><p><font face="Arial, Helvetica, sans-serif" size="-1"><b>Management Reporting</b><br>
			This section of the Online Management System provides a number of ways to:<br>
			<ul>
				<li>display the compliance ratings provided in audits, planning and compliance assessment activities, and;<br><br></li>
				<li>report on the service agreement activities - tasks that Safety &amp; Wellbeing has committed to providing to faculties and units over the lifetime of a Faculty/Unit Plan - across the University.</li>
			</ul><br><br>
			Select a report type from the list at left.</font></p>
</td>
        </tr>
        </table>
       
	</td>
</tr>
</table>

<!-- #Include file="include\footer.asp" -->