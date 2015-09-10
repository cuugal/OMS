<!-- #Include file="include\general.asp" -->
<%
	if SecurityCheck(2) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>
<% PageTitle = "Compliance Checking Report"%>
	
<!-- #Include file="include\header.asp" -->

<%
	dim con, ComplianceID
	dim sqlDate
	dim rsDate
	
	ComplianceID = request("ccID")
	
	set con = server.CreateObject("adodb.connection")
	con.Open "DSN=ehs"
	
	sqlDate = "SELECT ccActionPlan AS ActionPlan, ccDate AS CompDate, ccAssessor, apCompletionDate " & _
			  "FROM CC_Compliance INNER JOIN AP_ActionPlans ON CC_Compliance.ccActionPlan = AP_ActionPlans.apID " & _
			  "WHERE ccID = " & ComplianceID
	set rsDate = con.Execute(sqlDate)
%>
 <table width="100%" border="0" cellspacing="3">
  <tr> 
    <td><!-- removed old EHS Branch logo CL 14/4/09 <img src="ehslogo2.gif" width="142" height="111" alt="EHS logo" border="0">-->&nbsp;</td>
    <td> 
      <div align="right"><img src="utslogo.gif" width="135" height="30" alt="UTS logo" border="0"></div>
    </td>
  </tr>
  <tr> 
    <td colspan="2">&nbsp; 
      <table width="100%" border="1" cellspacing="1" cellpadding="0">
        <tr> 
          <td> 
            <table border="0" width="100%">
              <tr> 
                <td class="label" width="15%">Faculty / Unit:</td>
                <td><%=Session("DepName")%></td>
              </tr>
              <tr> 
                <td class="label">Name of Assessor:</td>
                <td><%=rsDate("ccAssessor")%></td>
              </tr>
              <tr> 
                <td class="label">Date:</td>
                <td><%=rsDate("CompDate")%></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <h2><br>
        &nbsp;&nbsp;STATUS OF COMPLIANCE WITH EHS PLAN</h2>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
			<p><FONT SIZE="1">*Compliance Ratings:<BR>
		0 = Non-compliant, 1 = Non-compliant - Some action evident but not yet compliant, 2 = Compliant - just requires maintenance, 3 = Best practice evident</FONT></p>
      <table width="100%" border="1">
        <tr> 
          <td class="label" width="33%">COMPLIANCE REQUIREMENTS</td>
          <td class="label" width="15%">RATING FROM LAST<br>
            EHS PLAN (<%=rsDate("apCompletionDate")%>)</td>
          <td class="label">NEW COMPLIANCE <br>
            RATING (0,1,2,3)</td>
        </tr><%	function ShowComplianceChecking()
		dim sqlSteps, sqlComp
		dim rsSteps, rsComp
		
		sqlSteps = "SELECT IN_Steps.stShortName, IN_Steps.stID " & _
				   "FROM IN_Steps " & _
				   "ORDER BY IN_Steps.stReportOrder"
		set rsSteps = con.Execute(sqlSteps)
%>	<%		
		while not rsSteps.eof			sqlComp = "SELECT irName, irId, arRating, cdNewRating " & _
					  "FROM IN_Requirements INNER JOIN ((CC_ComplianceDetails INNER JOIN CC_Compliance ON CC_ComplianceDetails.cdCompliance = CC_Compliance.ccID) INNER JOIN AP_Requirements ON (CC_Compliance.ccActionPlan = AP_Requirements.arActionPlan) AND (CC_ComplianceDetails.cdRequirement = AP_Requirements.arRequirement)) ON (CC_ComplianceDetails.cdRequirement = IN_Requirements.irId) AND (IN_Requirements.irId = AP_Requirements.arRequirement) " & _
					  "WHERE ccID = " & ComplianceID & " AND irStep = " & rsSteps("stID") & " and arSelected = Yes " & _
					  "ORDER BY irDisplayOrder"			set rsComp = con.Execute(sqlComp)			
%>
			<tr><td colspan="3"><br><b><%=rsSteps("stShortName")%></b><br><br></td></tr><%
			while not rsComp.eof
%>
				<tr>
					<td><%=rsComp("irName")%></td>
					<td>&nbsp;&nbsp;&nbsp;&nbsp;<%=rsComp("arRating")%></td>
					<td>&nbsp;&nbsp;&nbsp;&nbsp;<%=rsComp("cdNewRating")%></td>
				</tr><% rsComp.movenext
			wend
						rsSteps.movenext
		wend%><%
	end function
	
	ShowComplianceChecking%>        
      </table>
    </td>
  </tr>
</table>

<!-- #Include file="include\footer.asp" -->