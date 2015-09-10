<!-- #Include file="include\general.asp" -->
<%
	if SecurityCheck(1) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>
<% PageTitle = "Compliance Checking Worksheet"%>
	
<!-- #Include file="include\header.asp" -->

<%
	dim con, ActionPlan
	dim sqlDate
	dim rsDate
	
	ActionPlan = request("apID")
	
	set con = server.CreateObject("adodb.connection")
	con.Open "DSN=ehs"
			  
	sqlDate = "SELECT	apCompletionDate " & _
			  "FROM		AP_ActionPlans " & _
			  "WHERE	apID = " & ActionPlan
	set rsDate = con.Execute(sqlDate)
%>
<table width="100%" border="0" cellspacing="3">
  <tr> 
    <td><!-- commented out the old EHS Branch logo <img src="ehslogo2.gif" width="142" height="111" alt="EHS logo" border="0">-->&nbsp;</td>
    <td> 
      <div align="right"><img src="utslogo.gif" alt="UTS logo"></div>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> &nbsp; 
      <table width="100%" border="1" cellspacing="1" cellpadding="0">
        <tr> 
          <td> 
            <table border="0" width="100%">
              <tr> 
                <td class="label" width="15%">Faculty/Unit:</td>
                <td><%=Session("DepName")%></td>
              </tr>
              <tr> 
                <td class="label">Name of Assessor:</td>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td class="label">Date:</td>
                <td>&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="2"><h2>&nbsp;&nbsp;STATUS OF COMPLIANCE WITH HEALTH &amp; SAFETY PLAN</h2></td>
  </tr>

	<tr>
		<td colspan="2">&nbsp;&nbsp;*Compliance Ratings:<br>&nbsp;&nbsp;
		<b>0</b> = Non-Compliant;&nbsp;&nbsp;&nbsp;<b>1</b> = Non-Compliant - Some action evident but not yet compliant;&nbsp;&nbsp;&nbsp;<b>2</b> = Compliant - just requires maintenance;&nbsp;&nbsp;&nbsp;<b>3</b> = Best practice evident</td>
	</tr>

  <tr> 
    <td colspan="2"> 
      <table width="100%" border="1">
        <tr> 
          <td class="label" width="33%">COMPLIANCE REQUIREMENTS</td>
          <td class="label" width="15%">RATING FROM LAST<br>
            HEALTH AND SAFETY PLAN<br>(<%=rsDate("apCompletionDate")%>)</td>
          <td class="label">NEW COMPLIANCE RATING (0,1,2,3)</td>
        </tr><%	function ShowComplianceChecking()
		dim sqlSteps, sqlComp
		dim rsSteps, rsComp
		
		sqlSteps = "SELECT IN_Steps.stShortName, IN_Steps.stID " & _
				   "FROM IN_Steps " & _
				   "ORDER BY IN_Steps.stReportOrder"
		set rsSteps = con.Execute(sqlSteps)
%>	<%		
		while not rsSteps.eof
			sqlComp = "SELECT IN_Requirements.irName, AP_Requirements.arRating " & _
					  "FROM IN_Requirements INNER JOIN (AP_ActionPlans INNER JOIN AP_Requirements ON AP_ActionPlans.apID = AP_Requirements.arActionPlan) ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
					  "WHERE apID = " & ActionPlan & " AND arSelected = Yes AND irStep = " & rsSteps("stID")			set rsComp = con.Execute(sqlComp)			
%>
			<tr>
				<td colspan="3"><br><b><%=rsSteps("stShortName")%></b><br><br></td>
			</tr>
<%
			while not rsComp.eof
%>
				<tr>
					<td><%=rsComp("irName")%></td>
					<td>&nbsp;&nbsp;&nbsp;&nbsp;<%=rsComp("arRating")%></td>
					<td>&nbsp;</td>
				</tr><%				rsComp.movenext
			wend
						rsSteps.movenext
		wend%><%
	end function
	
	ShowComplianceChecking%>        
      </table>
    </td>
  </tr>
</table>
<!-- #Include file="include\footer.asp" -->