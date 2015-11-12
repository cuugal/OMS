<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(4) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
	
	dim reqID
	
	reqID = request("req")
	
	if reqID = "-1" then 
		Response.Redirect "ReportingMenu.asp"
	end if
%>


<% PageTitle = "EHS Online Management System!"%>
	
<!-- #Include file="include\header.asp" -->

<%
	dim con, startYear, endYear

	set con	= server.createobject ("adodb.connection")
		con.open "DSN=ehs"
		
	dim sqlReqName, sqlYear
	dim rsReqName, rsYear

	sqlReqName = "Select irName from IN_Requirements where irID = " & reqID
	set rsReqName = con.Execute(sqlReqName)
	
	' Determine the start and end years for this report
	sqlYear = "SELECT Min(apStartYear) AS minStart, Max(apStartYear) AS maxStart, datepart('yyyy', Min(ccDate)) AS minComp, datepart('yyyy', Max(ccDate)) AS maxComp " & _
			  "FROM CC_Compliance RIGHT JOIN AP_ActionPlans ON CC_Compliance.ccActionPlan = AP_ActionPlans.apID " & _
			  "WHERE apCompleted = Yes"
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
<table border="0" cellpadding="2" width=640 align="center">
		<tr><td width=90%></td><td valign=middle align=right><a href=http://www.uts.edu.au/><img src=utslogo.gif width=135 alt=UTS height=30 border="0"></a></td></tr>
</table>

<p>

<table
width=640
cellspacing=0
border=0
cellpadding=4><tr
bgcolor=0099cc>
      <td><font
size=+1
face=arial
color=white><b>Compliance Ratings for <%=rsReqName("irName")%></b></font></td>
      <td
align=right><table
cellpadding=4
cellspacing=0
border=0><tr>
            <td
><font
size=-1
face=arial>



              </font></td>
          </tr></table></td></tr></table>
  <table
cellspacing=4
border=0
cellpadding=4>
    <tr>
      <td
valign=top>
        <table
bgcolor=a0b8c8
width=100%
cellspacing=0
border=0
cellpadding="2">
          <tr>
            <td>
              <table
bgcolor=white
width=100%
cellspacing=0
border=0
cellpadding="2">
                <tr>
                  
				  <td align="center"> 
				  
				  
                    <table width="100%" cellspacing=1 border="0" cellpadding=3>
                      <tr bgcolor="#98B1CB"> 
                        <td><b>Faculty/ Unit</b></td>
<%
	dim curYear
	
	curYear = startYear

	while curYear <= endYear
%>
                        <td><b><%=curYear%></b></td>
<%
		curYear = curYear + 1
	wend
%>
                      </tr>
<%
	dim sqlDept
	dim rsDept

	sqlDept = "SELECT AD_Departments.dpName, AD_Departments.dpID " & _
			  "FROM AD_Departments " & _
			  "ORDER BY AD_Departments.dpName"
	set rsDept = con.Execute(sqlDept)
	
	while not rsDept.eof
%>

                      <tr> 
                        <td bgcolor="eeeeee"><%=rsDept("dpName")%></td>
<%
		curYear = startYear
		
		while curYear <= endYear
			sqlRate = "SELECT TOP 1 Rating from ( " & _
					  "SELECT Max(arRating) AS Rating, apCompletionDate as dateEntered " & _
					  "FROM AD_Departments INNER JOIN (AP_ActionPlans INNER JOIN AP_Requirements ON AP_ActionPlans.apID = AP_Requirements.arActionPlan) ON AD_Departments.dpID = AP_ActionPlans.apFaculty " & _
					  "WHERE arRequirement = " & reqID & " AND apCompleted = Yes AND arSelected = Yes AND dpID = " & rsDept("dpID") & " AND apStartYear = " & curYear & " " & _
					  "GROUP BY apCompletionDate " & _

					  "UNION  " & _

					  "SELECT Max([cdNewRating]) AS Rating,  ccDate as  dateEntered " & _
					  "FROM CC_ComplianceDetails INNER JOIN (CC_Compliance INNER JOIN (AD_Departments INNER JOIN AP_ActionPlans ON AD_Departments.dpID = AP_ActionPlans.apFaculty) ON CC_Compliance.ccActionPlan = AP_ActionPlans.apID) ON CC_ComplianceDetails.cdCompliance = CC_Compliance.ccID " & _
					  "WHERE cdRequirement = " & reqID & " AND dpID = " & rsDept("dpID") & " AND apStartYear = " & curYear & " " & _
					  "GROUP BY ccDate) " & _
					  "ORDER BY dateEntered desc, rating desc "
			set rsRate = con.Execute(sqlRate)
%>
                        <td bgcolor="eeeeee"><%if not rsRate.BOF then Response.Write rsRate("Rating")%></td>
<%
			curYear = curYear + 1
		wend
%>

                      </tr>
<%
		rsDept.movenext
		
		if not rsDept.eof then
%>
                      <tr> 
                        <td bgcolor="#C1CAE1"><%=rsDept("dpName")%></td>
<%
		curYear = startYear
		
		while curYear <= endYear
			sqlRate = "SELECT TOP 1 Rating from ( " & _
					  "SELECT Max(arRating) AS Rating, apCompletionDate as dateEntered " & _
					  "FROM AD_Departments INNER JOIN (AP_ActionPlans INNER JOIN AP_Requirements ON AP_ActionPlans.apID = AP_Requirements.arActionPlan) ON AD_Departments.dpID = AP_ActionPlans.apFaculty " & _
					  "WHERE arRequirement = " & reqID & " AND apCompleted = Yes AND arSelected = Yes AND dpID = " & rsDept("dpID") & " AND apStartYear = " & curYear & " " & _
					  "GROUP BY apCompletionDate " & _

					  "UNION  " & _

					  "SELECT Max([cdNewRating]) AS Rating,  ccDate as  dateEntered " & _
					  "FROM CC_ComplianceDetails INNER JOIN (CC_Compliance INNER JOIN (AD_Departments INNER JOIN AP_ActionPlans ON AD_Departments.dpID = AP_ActionPlans.apFaculty) ON CC_Compliance.ccActionPlan = AP_ActionPlans.apID) ON CC_ComplianceDetails.cdCompliance = CC_Compliance.ccID " & _
					  "WHERE cdRequirement = " & reqID & " AND dpID = " & rsDept("dpID") & " AND apStartYear = " & curYear & " " & _
					  "GROUP BY ccDate) " & _
					  "ORDER BY dateEntered desc, rating desc "
			set rsRate = con.Execute(sqlRate)
%>
                        <td bgcolor="#C1CAE1"><%if not rsRate.BOF then Response.Write rsRate("Rating")%></td>
<%
			curYear = curYear + 1
		wend
%>
                      </tr>
<%
			rsDept.movenext
		end if
	wend
%>
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




