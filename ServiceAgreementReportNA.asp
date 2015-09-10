<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(1) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>

<% PageTitle = "Service Agreement Report"%>

<!-- #Include file="include\header.asp" -->

<%
	dim con, ServiceAgreement, ActionPlan
	dim sqlAP
	dim rsAP

	set con	= server.createobject ("adodb.connection")	con.open "DSN=ehs"
	
	ServiceAgreement = Request("saID")
	ActionPlan = Request("apID")

		sqlAP = "SELECT * FROM AP_ActionPlans INNER JOIN AD_Departments ON AP_ActionPlans.apFaculty = AD_Departments.dpID WHERE apID = " & ActionPlan
	set rsAP = con.execute (sqlAP)
	
	' Get the requirements and their compliance ratings
	dim rsSA, sqlSA
			
	sqlSA =		"SELECT irID, irName as Requirement, arRating as Rating, sdEHSServices as EHSServices, sdContact as Contact, sdTimeFrame as TimeFrame, saAddEHSServices as AddEHSServices, saAddContact as AddContact, saAddTimeFrame as AddTimeFrame " & _
				"FROM SA_ServiceAgreement INNER JOIN (IN_Requirements INNER JOIN (SA_ServiceAgreementDetails INNER JOIN AP_Requirements ON SA_ServiceAgreementDetails.sdRequirement = AP_Requirements.arRequirement) ON IN_Requirements.irId = AP_Requirements.arRequirement) ON (SA_ServiceAgreement.saActionPlan = AP_Requirements.arActionPlan) AND (SA_ServiceAgreement.saID = SA_ServiceAgreementDetails.sdServiceAgreement) " & _
				"WHERE arSelected = Yes  AND saActionPlan = " & ActionPlan & " ORDER BY irDisplayOrder" 
				'saID = " & ServiceAgreement replaced with saActionPlan = " & ActionPlan
				'changed by DLJ 13feb4 to fix bug in display service agreement

	set rsSA = con.Execute (sqlSA)
%>
	
<table width="100%" border="0" cellspacing="3">
  <tr> 
    <td><!--<img src="ehslogo2.gif" width="142" height="111" alt="EHS logo" border="0">-->&nbsp;</td>
    <td> 
      <div align="right"><img src="utslogo.gif" width="135" height="30" alt="UTS logo graphic"></div>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <h2> <b>Health and Safety Service Agreement for the <%=rsAP("dpName")%></b></h2>
      The purpose of this Service Agreement is to document the services that the <%=rsAP("dpName")%> requires from Safety &amp; Wellbeing to assist in the implementation of its Health &amp; Safety Plan.<br><br>
      The health and safety planning and review cycle for the <%=rsAP("dpName")%> involves:
      <ul>
        <li>Updating the Health and Safety Plan every <%=rsAP("dpActionPlanDuration")%> years (with the next update to occur in <%=year(rsAP("apCompletionDate")) + rsAP("dpActionPlanDuration")%>)</li>
        <li>Reviewing the Health and Safety Plan and reporting on compliance at the mid-point 
          between plans (i.e. in <%=year(rsAP("apCompletionDate")) + cint(rsAP("dpActionPlanDuration") / 2)%>)</li>
      </ul>
      <h2>Health and Safety Service Agreement for the <%=rsAP("dpName")%> at <%=rsAP("apCompletionDate")%></h2>
    </td>
  </tr>
  <tr> 
    <td colspan="2">
    <p><font size="1">*Compliance Ratings:<BR>
		0 = Non-compliant, 1 = Non-compliant - Some action evident but not yet compliant, 2 = Compliant - just requires maintenance, 3 = Best practice evident</font></p>
      <table border="1" cellpadding="3">
        <tr> 
          <td class="label">COMPLIANCE REQUIREMENT</td>
          <td ><span class="label"><center>Compliance<br>
            rating at 
            <%=monthname(month(rsAP("apCompletionDate")))%>, <%=year(rsAP("apCompletionDate"))%>
            <br>
            </center></span></td>
          <td ><span class="label">HEALTH AND SAFETY SERVICES</span><br>
          </td>
          <td class="label">FACULTY/UNIT CONTACT</td>
          <td class="label">TIMEFRAME</td>
        </tr>
        <%

		
		while not rsSA.EOF
%>
        <tr> 
          <td> 
			<%=rsSA("Requirement")%>&nbsp;
          </td>
          <td> 
            <center><%=rsSA("Rating")%>&nbsp;</center>
          </td>
          <td> 
            <%=rsSA("EHSServices")%>&nbsp;
          </td>
          <td> 
            <%=rsSA("Contact")%>&nbsp;
          </td>
          <td> 
            <%=rsSA("TimeFrame")%>&nbsp;
          </td>
        </TR>
<%
			rsSA.MoveNext
		wend
		
		rsSA.movefirst
%>     
		<tr> 
          <td> 
			Additional Services
          </td>
          <td> 
            <center>--</center>
          </td>
          <td> 
            <%=rsSA("AddEHSServices")%>
          </td>
          <td> 
            <%=rsSA("AddContact")%>
          </td>
          <td> 
            <%=rsSA("AddTimeFrame")%>
          </td>
        </TR>
      </table>

      <table width="100%" border="0" cellspacing="0" id="table3">
        <tr>
          <td width="60%"><p>Signed</p><br><br><br></td>
          <td><p>Signed</p><br><br><br></td>
        </tr>
        <tr>
          <td><p>Safety &amp; Wellbeing Contact<br />Date</p></td>
          <td><p>Dean/Director<br />Date</p></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td colspan=2>
<%
	if request("draft") = "true" then
		Response.Write "<p align='center'>><input type=""submit"" value=""    Close Window    "" onclick=""window.close();"" id=submit1 name=submit1></p>"
	end if 
%>
    </td>
  </tr>
</table>

<!-- #Include file="include\footer.asp" -->