<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(4) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
	
	dim Year
	
	Year = request("year")
	
	if year = "-1" then 
		Response.Redirect "ReportingMenu.asp"
	end if
%>


<% PageTitle = "EHS Online Management System!"%>
	
<!-- #Include file="include\header.asp" -->

<%
	dim con

	set con	= server.createobject ("adodb.connection")
		con.open "DSN=ehs"
	
	dim sqlDep
	dim rsDep

	sqlDep = "SELECT dpName, dpID " & _
			  "FROM AD_Departments"
	set rsDep = con.Execute(sqlDep)
%>

<center> 
<table border="0" cellpadding="2" width="640" align="center">

      <tr>
        <td><BR><BR><BR></td>
      </tr>
      <tr>
        <td valign="middle" align="right">
          <a href="http://www.uts.edu.au/"><img src="utslogo.gif" width="135" height="30" alt="UTS" border="0"></a>
        </td>
      </tr>

</table>

<p>
<table cellspacing="0" border="0" cellpadding="4">
<tr bgcolor="#0099cc">
  <td><font size="+1" face="arial" color="white"><b>EHS Planning Session Compliance Ratings for <%=Year%></b></font></td>
  <td align="right">
    <table cellpadding="4" cellspacing="0" border="0">
      <tr>
        <td></td>
      </tr>
    </table>
  </td>
</tr>
</table>


<table cellspacing="0" border="0" cellpadding="4">
    <tr>
      <td valign="top">
        <table bgcolor="#a0b8c8" width="100%" cellspacing="0" border="0" cellpadding="2">
          <tr>
            <td>
              <table bgcolor="white" width="100%" cellspacing="0" border="0" cellpadding="2">
              <tr>
                <td align="center"> 
                    <table width="100%" cellspacing="1" border="0" cellpadding="3">
                      <tr> 
                        <td bgcolor="#98b1cb"><b>Compliance Requirement</b></td>
<%
	while not rsDep.EOF
%>
                        <td bgcolor="#98b1cb"><b><%=rsDep("dpName")%></b></td>
<%
		rsDep.movenext
	wend
%>
						<td bgcolor="#98b1cb"><b>Compliance Requirement</b></td>
                      </tr>
<%
	dim sqlReq, sqlRate
	dim rsReq, rsRate

	sqlReq =  "SELECT IN_Requirements.irName, IN_Requirements.irId " & _
			  "FROM IN_Requirements " & _
			  "ORDER BY IN_Requirements.irDisplayOrder"
	set rsReq = con.Execute(sqlReq)
	
	while not rsReq.eof
%>

                      <tr> 
                        <td bgcolor="#eeeeee"><%=rsReq("irName")%></td>
<%
		rsDep.movefirst
		
		while not rsDep.eof
			sqlRate = "SELECT TOP 1 Rating from ( " & _
					  "SELECT Max(arRating) AS Rating, apCompletionDate as dateEntered " & _
					  "FROM AD_Departments INNER JOIN (AP_ActionPlans INNER JOIN AP_Requirements ON AP_ActionPlans.apID = AP_Requirements.arActionPlan) ON AD_Departments.dpID = AP_ActionPlans.apFaculty " & _
					  "WHERE arRequirement = " & rsReq("irID") & " AND apCompleted = Yes AND arSelected = Yes AND dpID = " & rsDep("dpID") & " AND apStartYear = " & Year & " " & _
					  "group by apCompletionDate " & _

					  ") " & _
					  "order by dateEntered desc, rating desc "
			set rsRate = con.Execute(sqlRate)
%>
<!--   This SQL was removed from both the SQL queries above and below. From between the lines [ "group by apCompletionDate " & _ ]  and [ ") " & _ ] . Its removal is intended to remove inclusion of the Compliance Checking ratings, so that just the Planning Session scores are returned.
						"UNION  " & _

					  "SELECT Max([cdNewRating]) AS Rating,  ccDate as  dateEntered " & _
					  "FROM CC_ComplianceDetails INNER JOIN (CC_Compliance INNER JOIN (AD_Departments INNER JOIN AP_ActionPlans ON AD_Departments.dpID = AP_ActionPlans.apFaculty) ON CC_Compliance.ccActionPlan = AP_ActionPlans.apID) ON CC_ComplianceDetails.cdCompliance = CC_Compliance.ccID " & _
					  "WHERE cdRequirement = " & rsReq("irID") & " AND dpID = " & rsDep("dpID") & " AND apStartYear = " & Year & " " & _
					  "group by ccDate " & _ 
				-->
                        <td bgcolor="#eeeeee"><%if not rsRate.BOF then Response.Write rsRate("Rating")%></td>
<%
			rsDep.movenext
		wend
%>
						<td bgcolor="#eeeeee"><%=rsReq("irName")%></td>
                      </tr>
<%
		rsReq.movenext
		
		if not rsReq.eof then
%>
                      <tr> 
                        <td bgcolor="#C1CAE1"><%=rsReq("irName")%></td>
<%
			rsDep.movefirst
			
			while not rsDep.eof
				sqlRate = "SELECT TOP 1 Rating from ( " & _
					  "SELECT Max(arRating) AS Rating, apCompletionDate as dateEntered " & _
					  "FROM AD_Departments INNER JOIN (AP_ActionPlans INNER JOIN AP_Requirements ON AP_ActionPlans.apID = AP_Requirements.arActionPlan) ON AD_Departments.dpID = AP_ActionPlans.apFaculty " & _
					  "WHERE arRequirement = " & rsReq("irID") & " AND apCompleted = Yes AND arSelected = Yes AND dpID = " & rsDep("dpID") & " AND apStartYear = " & Year & " " & _
					  "group by apCompletionDate " & _

					  ") " & _
					  "order by dateEntered desc, rating desc "
			set rsRate = con.Execute(sqlRate)
%>
	                        <td bgcolor="#C1CAE1"><%if not rsRate.BOF then Response.Write rsRate("Rating")%></td>
<%
				rsDep.movenext
			wend
%>
						<td bgcolor="#C1CAE1"><%=rsReq("irName")%></td>
                      </tr>
<%
			rsReq.movenext
		end if
	wend
%>
					<tr> 
                        <td bgcolor="#98b1cb"><b>Compliance Requirement</b></td>
<%
	rsDep.movefirst

	while not rsDep.EOF
%>
                        <td bgcolor="#98b1cb"><b><%=rsDep("dpName")%></b></td>
<%
		rsDep.movenext
	wend
%>
						<td bgcolor="#98b1cb"><b>Compliance Requirement</b></td>
                      </tr>

                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
        <br>
      </td>
    </tr>
  </table>

</center></body></html>




