<!-- #Include file="include\general.asp" -->

<%
	if SecurityCheck(3) = false then ' User must have write access for this department
		Response.Redirect ("restricted.asp")
		Response.end
	end if
%>

<% PageTitle = "Service Agreement Form"%>

<!-- #Include file="include\header.asp" -->

<%
	dim con, ActionPlan, ServiceAgreement
	dim sqlAP
	dim rsAP

	set con	= server.createobject ("adodb.connection")	con.open "DSN=ehs"
	
	ActionPlan = Request("apID")
	ServiceAgreement = Request("saID")

	sqlAP = "Select * from AP_ActionPlans, AD_Departments where apID = " & ActionPlan
	set rsAP = con.execute (sqlAP)
	
	' DELETE ANY DRAFT OR EXISTING SA FOR THE CURRENT DEPARTMENT
	if request("exists") = "yes" then
		' This query will do a cascade delete and also delete the ServiceAgreementDetails
		sqlDel = "DELETE SA_ServiceAgreement.saActionPlan " & _
				 "FROM SA_ServiceAgreement " & _
				 "WHERE saActionPlan = " & ActionPlan
		con.Execute (sqlDel) 
		
		RefreshParent()
	end if
	
	' Get the requirements and their compliance ratings
	dim rsSA, sqlSA
	
	if ServiceAgreement = "" then
		sqlSA = "SELECT irID, irName as Requirement, arRating as Rating, '' as EHSServices, '' as Contact, '' as TimeFrame, '' as AddEHSServices, '' as AddContact, '' as AddTimeFrame " & _
				"FROM IN_Requirements INNER JOIN AP_Requirements ON IN_Requirements.irId = AP_Requirements.arRequirement " & _
				"WHERE arSelected = Yes AND arActionPlan =  " & ActionPlan & " " & _
				"ORDER BY irDisplayOrder"
	else		
		sqlSA = "SELECT irID, irName as Requirement, arRating as Rating, sdEHSServices as EHSServices, sdContact as Contact, sdTimeFrame as TimeFrame, saAddEHSServices as AddEHSServices, saAddContact as AddContact, saAddTimeFrame as AddTimeFrame " & _
				"FROM SA_ServiceAgreement INNER JOIN (IN_Requirements INNER JOIN (SA_ServiceAgreementDetails INNER JOIN AP_Requirements ON SA_ServiceAgreementDetails.sdRequirement = AP_Requirements.arRequirement) ON IN_Requirements.irId = AP_Requirements.arRequirement) ON (SA_ServiceAgreement.saActionPlan = AP_Requirements.arActionPlan) AND (SA_ServiceAgreement.saID = SA_ServiceAgreementDetails.sdServiceAgreement) " & _
				"WHERE arSelected = Yes AND arActionPlan =  " & ActionPlan & " " & _
				"ORDER BY irDisplayOrder"
	end if	
	
	set rsSA = con.Execute (sqlSA)
%>
	

<table width="100%" border="0" cellspacing="3">
  <tr> 
    <td></td>
    <td> 
      <div align="right"><img src="utslogo.gif" width="135" height="30"></div>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
		<div><h2>Draft Health and Safety Service Agreement for the <%=Session("DepName")%></h2></div>
    </td>
</tr>
  <!--<tr> 
    <td colspan="2"  class="label"> 
      <b>Service Agreement for the <%=Session("DepName")%></b><br><br>
    </td>
  </tr>-->
  <tr> 
    <td colspan="2">
      
      <form action="ServiceAgreementForm_Process.asp" method="post" name="saForm"> 
      <input type="hidden" name="action" value="">
	  <input type="hidden" name="apID" value="<%=ActionPlan%>">
	  <input type="hidden" name="saID" value="<%=ServiceAgreement%>">
      
      <table border="1" cellpadding="3">
        <tr> 
          <td class="label">Compliance Requirement</td>
          <td><center><b>Compliance<br>rating at<br> 
            <%=monthname(month(rsAP("apCompletionDate")))%>, <%=year(rsAP("apCompletionDate"))%>
            </center></b></td>
          <td><span class="label">Health and Safety Services   <i>(max. 250 characters)</i></span><br>
          </td>
          <td class="label">Faculty/Unit Contact</td>
          <td class="label">Timeframe</td>
        </tr>
        <%

		
		while not rsSA.EOF
%>
        <tr> 
          <td>
			<%=Response.Write(rsSA("Requirement"))%>
          </td>
          <td> 
            <center><%=rsSA("Rating")%></center>
          </td>
          <td> 
			<textarea name="serv_<%=rsSA("irID")%>"  rows=5 cols=50 wrap=virtual><%=rsSA("EHSServices")%></textarea>
		  </td>
          <td> 
            <INPUT type="text" name="cont_<%=rsSA("irID")%>" value="<%=rsSA("Contact")%>" size="20" maxlength=100>
          </td>
          <td> 
            <INPUT type="text" name="time_<%=rsSA("irID")%>" value="<%=rsSA("TimeFrame")%>" size="20" maxlength=100>
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
            <textarea name="serv_ADD" rows=5 cols=50  wrap=virtual><%=rsSA("AddEHSServices")%></textarea>
          </td>
          <td> 
            <input type="text" name="cont_ADD" value="<%=rsSA("AddContact")%>" size="20" maxlength=100>
          </td>
          <td> 
            <input type="text" name="time_ADD" value="<%=rsSA("AddTimeFrame")%>" size="20" maxlength=100>
          </td>
        </TR>
      </table>
      
      <P> 
	      <input type="submit" value="    Save as Draft    " onclick="javascript:saForm.action.value='draft'">
			&nbsp;&nbsp;&nbsp;&nbsp;
		  <input type="submit" value="    Save as Final    " onclick="javascript:saForm.action.value='final'">
	  </P>
		
      </form>
      
      <P>&nbsp;</P>
    </td>
  </tr>
</table>

</form>

<!-- #Include file="include\footer.asp" -->