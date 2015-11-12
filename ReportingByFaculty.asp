<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(2) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
	
	dim facID
	
	facID = request("fac")
	
	if facID = "-1" then 
		Response.Redirect "ComplianceMenu.asp"
	end if
%>


<% PageTitle = "EHS Online Management System!"%>
	
<!-- #Include file="include\header.asp" -->

<%
	dim con, startYear, endYear

	set con	= server.createobject ("adodb.connection")
		con.open "DSN=ehs"
	
	' Determine the start and end years for this report
	sqlYear = "SELECT Min(apStartYear) AS minStart, Max(apStartYear) AS maxStart, datepart('yyyy', Min(ccDate)) AS minComp, datepart('yyyy', Max(ccDate)) AS maxComp " & _
			  "FROM CC_Compliance RIGHT JOIN AP_ActionPlans ON CC_Compliance.ccActionPlan = AP_ActionPlans.apID " & _
			  "WHERE apFaculty = " & facID & " AND apCompleted = Yes"
	set rsYear = con.Execute(sqlYear)

	
	
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
%>

<center> 
<table border="0" cellpadding="2" width="640" align="center">
<tr>
	<td width="1%">

		<table>
			<tr>
				<td></td>
			</tr>
		</table>

	</td>

	<td></td>

	<td width="1%">

		<table>
			<tr>
				<td><br><br><br></td>
			</tr>

			<tr>
				<td valign="middle" align="right"><a href="http://www.uts.edu.au/" title="UTS Home"><img src="utslogo.gif" width="135" alt="UTS" height="30" border="0"></a></td>
			</tr>
		</table>

	</td>
</tr>
</table>

<table width="640" cellspacing="0" border="0" cellpadding="4">
	<tr bgcolor="#0099cc">
		<td>
			<font size="+1" face="Arial" color="white">
				<b>Compliance Ratings for <%=Session("DepName")%></b>
			</font>
		</td>
		
		<td>&nbsp;</td>
	</tr>

<tr>
<td><table width="640" cellspacing="1" cellpadding="3" style="border: solid 2px #a0b8c8; padding: 2px; ">
	<tr>
		<td bgcolor="#98b1cb"><b>Compliance Requirement</b></td>
<%
	dim curYear
	
	curYear = startYear

	while curYear <= endYear
%>
     <td bgcolor="#98b1cb"><center><b><%=curYear%></b></center></td>
<%
		curYear = curYear + 1
	wend
%>
                      </tr>
<%
	dim sqlReq, sqlRate
	dim rsReq, rsRate

	sqlReq =  "SELECT IN_Requirements.irName, IN_Requirements.irId " & _
			  "FROM IN_Requirements " & _
			  "WHERE irActive = Yes " & _
			  "ORDER BY IN_Requirements.irDisplayOrder"
	set rsReq = con.Execute(sqlReq)
	
	while not rsReq.eof
%>

                      <tr> 
                        <td bgcolor="#eeeeee"><%=rsReq("irName")%></td>
<%
		curYear = startYear
		
		while curYear <= endYear
			sqlRate = "SELECT TOP 1 Rating from ( " & _
					  "SELECT Max(arRating) AS Rating, apCompletionDate as dateEntered " & _
					  "FROM AD_Departments INNER JOIN (AP_ActionPlans INNER JOIN AP_Requirements ON AP_ActionPlans.apID = AP_Requirements.arActionPlan) ON AD_Departments.dpID = AP_ActionPlans.apFaculty " & _
					  "WHERE arRequirement = " & rsReq("irID") & " AND apCompleted = Yes AND arSelected = Yes AND dpID = " & facID & " AND apStartYear = " & curYear & " " & _
					  "group by apCompletionDate " & _

					  "UNION  " & _

					  "SELECT Max([cdNewRating]) AS Rating,  ccDate as  dateEntered " & _
					  "FROM CC_ComplianceDetails INNER JOIN (CC_Compliance INNER JOIN (AD_Departments INNER JOIN AP_ActionPlans ON AD_Departments.dpID = AP_ActionPlans.apFaculty) ON CC_Compliance.ccActionPlan = AP_ActionPlans.apID) ON CC_ComplianceDetails.cdCompliance = CC_Compliance.ccID " & _
					  "WHERE cdRequirement = " & rsReq("irID") & " AND dpID = " & facID & " AND apStartYear = " & curYear & " " & _
					  "group by ccDate) " & _
					  "order by dateEntered desc, rating desc "
			set rsRate = con.Execute(sqlRate)
%>
                        <td bgcolor="#eeeeee"><center><%if not rsRate.BOF then Response.Write rsRate("Rating")%></center></td>
<%
			curYear = curYear + 1
		wend
%>

                      </tr>
<%
		rsReq.movenext
		
		if not rsReq.eof then
%>
            <tr> 
            <td bgcolor="#c1cae1"><%=rsReq("irName")%></td>
<%
			curYear = startYear
			
			while curYear <= endYear
				sqlRate = "SELECT TOP 1 Rating from ( " & _
					  "SELECT Max(arRating) AS Rating, apCompletionDate as dateEntered " & _
					  "FROM AD_Departments INNER JOIN (AP_ActionPlans INNER JOIN AP_Requirements ON AP_ActionPlans.apID = AP_Requirements.arActionPlan) ON AD_Departments.dpID = AP_ActionPlans.apFaculty " & _
					  "WHERE arRequirement = " & rsReq("irID") & " AND apCompleted = Yes AND arSelected = Yes AND dpID = " & facID & " AND apStartYear = " & curYear & " " & _
					  "group by apCompletionDate " & _

					  "UNION  " & _

					  "SELECT Max([cdNewRating]) AS Rating,  ccDate as  dateEntered " & _
					  "FROM CC_ComplianceDetails INNER JOIN (CC_Compliance INNER JOIN (AD_Departments INNER JOIN AP_ActionPlans ON AD_Departments.dpID = AP_ActionPlans.apFaculty) ON CC_Compliance.ccActionPlan = AP_ActionPlans.apID) ON CC_ComplianceDetails.cdCompliance = CC_Compliance.ccID " & _
					  "WHERE cdRequirement = " & rsReq("irID") & " AND dpID = " & facID & " AND apStartYear = " & curYear & " " & _
					  "group by ccDate) " & _
					  "order by dateEntered desc, rating desc "
				set rsRate = con.Execute(sqlRate)
%>
	                        <td bgcolor="#c1cae1"><center><%if not rsRate.BOF then Response.Write rsRate("Rating")%></center></td>
<%
				curYear = curYear + 1
			wend
%>
                      </tr>
<%
			rsReq.movenext
		end if
	wend
%>
     </table>
</td>
</tr>
</table>
</center>

</body>
</html>