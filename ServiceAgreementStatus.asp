<%@language = VBscript%>
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
    dim numRecordCounter
    dim strSQL 
    dim strSQL1
    dim rsAdd
    dim rsADDSA
    dim numRecords
    dim StrServiceAction
    dim numReqId
    dim strComments
    dim SA
    dim C
    dim AP
    dim servAgr
'*******************************Database connectivity code************************************************    

	set con	= server.createobject ("adodb.connection")
	con.open "DSN=ehs"
	
'**************************code to edit the form *********************************************************
    
if Request.Form("btnSave")="Save Form" then  
    
'*********************gathering information from the existing form***************************************
 numRecords = Request.Form("hdnRecordCount")
 
'********************************************************************************************************
   SA = Request.Form("chkSA")
   C = Request.Form("txtASComments")
   AP= Request.Form("hdnsaActionPlan")
   servAgr = Request.Form("hdnServiceAgreement")
 '  response.write(servAgr)
   
'***********************************Applying a loop******************************************************

For numRecordCounter = 1 to numRecords

   strServiceAction = Request.Form("chkServiceAction" + cstr(numRecordCounter))
   numReqId = Request.Form("hdnSdRequirement" + cstr(numRecordCounter))
   strComments = Request.Form("txtComments" + cstr(numRecordCounter))
 
   
  ' response.write(numReqId)%><%
  'response.write(strServiceAction)
  
   if strServiceAction = "on" then %><%
          strServiceAction = "Yes"
          'response.write(strServiceAction)
          'response.write(numReqId) %><%

          else
          strServiceAction = "No"
          'response.write(strServiceAction)%><%

     end if
   strSQL = "UPDATE SA_ServiceAgreementDetails SET sdServiceActioned = "&strServiceAction&",sdComments ='"&strComments&"' where sdServiceAgreement = "&servAgr&" and sdRequirement ="&numReqId 


'************************************loop ends here******************************************************	
set rsAdd = con.Execute (strSQL)

next
      if SA = "on" then %><%
          sA = "Yes"
          else
          SA = "No"
      end if
   strSQL1 = "UPDATE SA_ServiceAgreement SET saSA = "&SA&",saC ='"&C&"' where saActionPlan = "&AP
  set rsAddSA = con.Execute (strSQL1)
Response.Write ("The Service Agreement has been Updated")
Response.end 

end if 
'********************************************************************************************************

	ServiceAgreement = Request("saID")
	ActionPlan = Request("apID")

		sqlAP = "SELECT * FROM AP_ActionPlans INNER JOIN AD_Departments ON AP_ActionPlans.apFaculty = AD_Departments.dpID WHERE apID = " & ActionPlan
	set rsAP = con.execute (sqlAP)
	
	' Get the requirements and their compliance ratings
	dim rsSA, sqlSA
			
	sqlSA =		"SELECT irID, irName as Requirement, arRating as Rating,sdID,saSA,saC,saActionPlan,sdServiceActioned,sdRequirement,sdComments,sdEHSServices as EHSServices, sdContact as Contact, sdTimeFrame as TimeFrame, saAddEHSServices as AddEHSServices, saAddContact as AddContact, saAddTimeFrame as AddTimeFrame " & _
				"FROM SA_ServiceAgreement INNER JOIN (IN_Requirements INNER JOIN (SA_ServiceAgreementDetails INNER JOIN AP_Requirements ON SA_ServiceAgreementDetails.sdRequirement = AP_Requirements.arRequirement) ON IN_Requirements.irId = AP_Requirements.arRequirement) ON (SA_ServiceAgreement.saActionPlan = AP_Requirements.arActionPlan) AND (SA_ServiceAgreement.saID = SA_ServiceAgreementDetails.sdServiceAgreement) " & _
				"WHERE arSelected = Yes  AND saActionPlan = " & ActionPlan & " ORDER BY irDisplayOrder" 
				'saID = " & ServiceAgreement replaced with saActionPlan = " & ActionPlan
				'changed by DLJ 13feb4 to fix bug in display service agreement

	set rsSA = con.Execute (sqlSA)
%>
	
<form name ="audit" action="SARP.asp" method="POST">
<table width="100%" border="0" cellspacing="3" id="table1">
  <tr> 
    <td><!-- <img src="ehslogo2.gif" width="142" height="111" alt="EHS logo" border="0"> removed by CL/Safety and Wellbeing 5/10/2010-->&nbsp;</td>
    <td> 
      <div align="right"><img src="utslogo.gif" width="135" height="30"></div>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <h2><b>Health and Safety Service Agreement Status for the <%=rsAP("dpName")%> as at <% response.write FormatDateTime(Date, 1) %>
			</b></h2>
      This Service Agreement documents the services that the <%=rsAP("dpName")%> requires from Safety &amp; Wellbeing to assist in the implementation of its Health and Safety Plan.<br><br>
      The health and safety management system cycle for the <%=rsAP("dpName")%> involves:
      <ul>
        <li>Updating the Health and Safety Plan every <%=rsAP("dpActionPlanDuration")%> years (with the next update to occur in <%=year(rsAP("apCompletionDate")) + rsAP("dpActionPlanDuration")%>)</li>
        <li>Reviewing the Health and Safety Plan and reporting on compliance at the mid-point between Plans (i.e. in <%=year(rsAP("apCompletionDate")) + cint(rsAP("dpActionPlanDuration") / 2)%>)</li>
      </ul>
     </td>
  </tr>
  <tr> 
    <td colspan="2">
		<font size="1">*Compliance Ratings:<br>
			0 = Non-compliant, 1 = Non-compliant - Some action evident but not yet compliant, 2 = Compliant - just requires maintenance, 3 = Best practice evident</font>
			<table border="1" cellpadding="3" id="table2">
        <tr> 
          <td class="label">COMPLIANCE REQUIREMENT</td>
          <td><span class="label"><center>Compliance<br>
            rating at 
            <%=monthname(month(rsAP("apCompletionDate")))%>, <%=year(rsAP("apCompletionDate"))%>
            <br>
            </center></span> </td>
          <td><span class="label">HEALTH AND SAFETY SERVICES</span><br></td>
          <td class="label">FACULTY/UNIT CONTACT</td>
          <td class="label">TIMEFRAME</td>
          <td class="label">SERVICE ACTIONED</td>
          <td class="label">COMMENTS BY SAFETY &amp; WELLBEING</td>
        </tr>
        <%
       
		numRecordCounter = 0
		while not rsSA.EOF
    	'***********************************checking the blank records for the service agreement*************************
		  if rsSA("EHSServices") <>"" then
		'****************************************************************************************************************
		dim boolSA 
		dim sdRequirement
		 sdRequirement = rsSA("sdRequirement")
		 numRecordCounter = numRecordCounter + 1
		boolSA = rsSA("sdServiceActioned")
		'response.write(boolSA)

		%>
        <tr>
        <input type="hidden" name="hdnSdRequirement<%=numRecordCounter%>" value="<%=sdRequirement%>">
          <td><%=rsSA("Requirement")%>&nbsp;</td>
          <td><center><%=rsSA("Rating")%>&nbsp;</center></td>
          <td><%=rsSA("EHSServices")%>&nbsp;</td>
          <td><%=rsSA("Contact")%>&nbsp;</td>
          <td><%=rsSA("TimeFrame")%>&nbsp;</td>
           <td><center>
						<input type="checkbox" name=chkServiceAction<%=numRecordCounter%>
				<% if rsSA("sdServiceActioned") = "True" then%> CHECKED <%end if%>></td>
          <td><%=rsSA("sdComments")%>
						</center></td>
        </tr>
<%              
                  end if
			rsSA.MoveNext
		wend
		 
		rsSA.movefirst
%>     
   <%  if rsSA("AddEHSServices")<>"" then %>
		<tr> 
          <td>Additional Services</td>
          <td><center>--</center></td>
          <td><%=rsSA("AddEHSServices")%></td>
          <td><%=rsSA("AddContact")%></td>
          <td><%=rsSA("AddTimeFrame")%></td>
          <td><center>
						<input type="checkbox" name=chkSA <%if rsSA("saSA") ="True" then%> checked <%end if%>></center></td>
         <td><%=rsSA("saC")%><%end if%></td>
        </tr>
      </table>

    </td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
  <input type="hidden" name="hdnRecordCount" value="<%=numRecordCounter%>">
  <input type="hidden" name="hdnsaActionPlan" value="<%=ActionPlan%>">
  <input type="hidden" name="hdnServiceAgreement" value="<%=ServiceAgreement%>">
</table>

</form>

<!-- #Include file="include\footer.asp" -->