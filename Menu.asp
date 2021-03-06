<!-- #Include file="include\general.asp" -->

<% 
  PageTitle = "Health and Safety Online Management System!" 
  PageName  = "menu.asp"
%>
  
<!-- #Include file="include\header_menu.asp" -->
<!--#include file="adovbs.inc"--> 
<%
  dim con
  dim sqlDraft, sqlFinal, sqlDepList, sqlAPYears
  dim rsDraft, rsFinal, rsDepList, rsAPYears

  Set objCmd  = Server.CreateObject("ADODB.Command")	
  objCmd.CommandType = adCmdText
 
  set con = server.createobject ("adodb.connection")
  con.open "DSN=ehs"
  
   Set objCmd.ActiveConnection = con	
  

	'AA escaping inputs via parameters
  'sqlDraft = "Select * from AP_ActionPlans where apFaculty = " & Session("DepID") & " and apCompleted = 0"
  sqlDraft = "Select * from AP_ActionPlans where apFaculty = ? and apCompleted = 0"
  objCmd.CommandText = sqlDraft
  objCmd.Parameters.Append objCmd.CreateParameter("apFaculty", adWChar, adParamInput, 50)
  objCmd.Parameters("apFaculty") = Session("DepID")
  'set rsDraft = con.execute(sqlDraft)
  set rsDraft =  server.createobject("adodb.recordset")
  rsDraft.Open objCmd

  'Rather than delete paramaters, just reinstantiate the entire cmd object
  Set objCmd  = Server.CreateObject("ADODB.Command")
  objCmd.CommandType = adCmdText
  Set objCmd.ActiveConnection = con
  
  'AA escaping inputs via parameters
  'sqlFinal = "Select * from AP_ActionPlans where apFaculty = " & Session("DepID") & " and apCompleted = 1"
  sqlFinal = "Select * from AP_ActionPlans where apFaculty = ? and apCompleted = 1"
  objCmd.CommandText = sqlFinal
  objCmd.Parameters.Append objCmd.CreateParameter("apFaculty", adWChar, adParamInput, 50)
  objCmd.Parameters("apFaculty") = Session("DepID")
  'set rsFinal = con.execute(sqlDraft)
  set rsFinal =  server.createobject("adodb.recordset")
  rsFinal.Open objCmd
  
  Set objCmd  = Server.CreateObject("ADODB.Command")
  objCmd.CommandType = adCmdText
  Set objCmd.ActiveConnection = con
  ' AA escaping inputs via params
  'sqlAPYears = "Select dpActionPlanDuration from AD_Departments where dpID = " & Session("DepID")
  sqlAPYears = "Select dpActionPlanDuration from AD_Departments where dpID = ?"
  objCmd.CommandText = sqlAPYears
  objCmd.Parameters.Append objCmd.CreateParameter("dpID", adWChar, adParamInput, 50)
  objCmd.Parameters("dpID") = Session("DepID")
  'set rsAPYears = con.Execute(sqlAPYears)
  set rsAPYears =  server.createobject("adodb.recordset")
  rsAPYears.Open objCmd
  
  
  
  
  Set objCmd  = Server.CreateObject("ADODB.Command")
  objCmd.CommandType = adCmdText
  Set objCmd.ActiveConnection = con
  
  ' Get the draft ServiceAgreement
  ' AA escaping inputs via params
  sqlSADraft = "SELECT apID, apStartYear, saID " & _
         "FROM AP_ActionPlans INNER JOIN SA_ServiceAgreement ON AP_ActionPlans.apID = SA_ServiceAgreement.saActionPlan " & _
		 "WHERE apFaculty = ? AND saComplete = No"
  ' "WHERE apFaculty = " & Session("DepID") & " AND saComplete = No"
  objCmd.CommandText = sqlSADraft
  objCmd.Parameters.Append objCmd.CreateParameter("dpID", adWChar, adParamInput, 50)
  objCmd.Parameters("dpID") = Session("DepID")
  'set rsSADraft = con.execute(sqlSADraft)
  set rsSADraft =  server.createobject("adodb.recordset")
  rsSADraft.Open objCmd
  
  'Response.write (rsSADraft)
  
  Set objCmd  = Server.CreateObject("ADODB.Command")
  objCmd.CommandType = adCmdText
  Set objCmd.ActiveConnection = con

  ' Get the final ServiceAgreements
  sqlSAFinal = "SELECT apID, apStartYear,saID " & _
         "FROM AP_ActionPlans INNER JOIN SA_ServiceAgreement ON AP_ActionPlans.apID = SA_ServiceAgreement.saActionPlan " & _
         "WHERE apFaculty = ? AND saComplete = No"
  ' "WHERE apFaculty = " & Session("DepID") & " AND saComplete = No"
  objCmd.CommandText = sqlSAFinal
  objCmd.Parameters.Append objCmd.CreateParameter("dpID", adWChar, adParamInput, 50)
  objCmd.Parameters("dpID") = Session("DepID")
  'set rsSAFinal = con.execute(sqlSAFinal)
  set rsSAFinal =  server.createobject("adodb.recordset")
  rsSAFinal.Open objCmd
  
  
  Set objCmd  = Server.CreateObject("ADODB.Command")
  objCmd.CommandType = adCmdText
  Set objCmd.ActiveConnection = con
  
  ' Get the ID of the ActionPlan that the FA and CC worksheets will be attached to
  sqlAPID = "select apID " & _
        "from AP_ActionPlans " & _
        "where apFaculty = ? and apStartYear = ( " & _
        "   SELECT max(apStartYear) " & _
        "   FROM AP_ActionPlans " & _
        "   WHERE apFaculty = ? and apCompleted = Yes )"
  objCmd.CommandText = sqlAPID
  objCmd.Parameters.Append objCmd.CreateParameter("dpID", adWChar, adParamInput, 50)
  objCmd.Parameters("dpID") = Session("DepID") 
  objCmd.Parameters.Append objCmd.CreateParameter("dpID1", adWChar, adParamInput, 50)
  objCmd.Parameters("dpID1") = Session("DepID")
  'set rsAPID = con.Execute(sqlAPID) 
  set rsAPID =  server.createobject("adodb.recordset")
  rsAPID.Open objCmd
  
  if not rsAPID.BOF then
    ActionPlan = rsAPID("apID")
  else
    ActionPlan = ""
  end if
%>

<table width="970" cellspacing="0" border="0" cellpadding="4" align="center">
<tr bgcolor="#0f4beb">
  <td valign="middle">
    <font size="+1" face="Arial" color="white"><b>&nbsp;<% =Session("DepName") %></b></font>
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<!--  <font size="-1" face="arial" color="white"><%=Session("AccessLevel")%></font> -->
  </td>
    <td align="right">
<%
  if SecurityCheck(4) = true and SecurityCheck(5) = false then
    sqlDepList = "Select dpID, dpName from AD_Departments ORDER by dpName"
    set rsDepList = con.Execute (sqlDepList)
%>    
    <table cellpadding="4" cellspacing="0" border="0">
    
    <form action="Menu.asp" method="post" name="Menuform">
    <tr>
      <td>
        <font size="-1" face="Arial">
        <select name="department" onchange="javascript:this.form.submit();">
          <option value="">- Please Select a Department/Faculty -</option>
<%
      while not rsDepList.EOF
%>
          <option value=<%=rsDepList("dpid")%> <%if cint(rsDepList("dpid")) = cint(Session("DepID")) then Response.Write " selected"%>><%=rsDepList("dpname")%></option>
<%
        rsDepList.MoveNext
      wend
%>
        </select>
        </font>
        
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       <font size="-1" face="arial" color="white"><%=Session("AccessLevel")%></font>
      </td> 
    </tr>
    
    </form>
    
    </table>
<%
  end if
%>         
  </td>
</tr>
</table>
  
<table width="970" cellspacing="4" border="0" cellpadding="4" align="center">
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
                <font size="2" face="arial" color="#3366cc"><b>Choose from the following areas:</b></font>
              </td>
            </tr>
            <tr bgcolor="#eeeeee"> 
              <td bgcolor="#eeeeee">
                <font size="2" face="arial" color="#333333">
                <a href="ActionPlanMenu.asp">Planning</a></font>
              </td>
            </tr>
<%
            if SecurityCheck(1) = true then ' dlj CHANGED SECURITY LEVEL FROM 2
              ' if rsAPYears("dpActionPlanDuration") = 2 then  ' DLJ commented this IF STATEMENT out - dont know why it was included - only allows viewing by high-riskers
%>
                <!--tr bgcolor="#eeeeee"> 
                  <td bgcolor="#eeeeee">
                  <font size="2" face="Arial" color="#333333"><a href="ServiceAgreementMenu.asp">Service Agreements</a></font>
                  </td>
                </tr-->
<%
              'end if
            end if
%>
<%
            if SecurityCheck(4) = true then
%>
            <tr bgcolor="#eeeeee"> 
              <td bgcolor="#eeeeee">
                <font size="2" face="Arial" color="#333333"><a href="AuditMenu.asp">Audits</a></font>
              </td>
            </tr>
			<%
            end if
			%>
            <!-- tr bgcolor="#eeeeee"> 
              <td bgcolor="#eeeeee">
                <font size="2" face="Arial" color="#333333"><a href="ComplianceMenu.asp">Compliance Assessment</a></font>
              </td>
            </tr -->
<%
            if SecurityCheck(4) = true then
%>
            <tr bgcolor="#eeeeee"> 
              <td bgcolor="#eeeeee">
                <font size="2" face="Arial" color="#333333"><a href="ReportingMenu.asp">Management Reporting</a></font>
              </td>
            </tr>
<%
            end if

            if SecurityCheck(3) = true then
%>
            <tr bgcolor="#eeeeee"> 
              <td bgcolor="#eeeeee"> 
                <p><font face="Arial" size="2" color="333333"><a href="PasswordAdmin.asp">Manage Passwords</a></font> 
              </td>
            </tr>
<%
            end if
%>
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
           <table  width="100%" cellspacing="6" border="0" cellpadding="6">
            <tr bgcolor="#eeeeee"> 
              <td valign="top">
                <font face="Arial"><b>Need more information?</b></font>
              </td>
            </tr>
            <tr bgcolor="#eeeeee"> 
              <td><font face="Arial, Helvetica, sans-serif" size="-1">Need more information on the UTS Health and Safety Management System?<br><br>An outline of the system is available from the <a href="http://www.safetyandwellbeing.uts.edu.au/">Safety &amp; Wellbeing Branch web site</a>.</font></td>
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

    <table bgcolor="#a0b8c8" width="100%" cellspacing="0" border="0" cellpadding="2">
    <tr>
      <td>
      
        <table bgcolor="white" width="100%" cellspacing="0" border="0" cellpadding="2">
        <tr>
          <td align="center">
                    
            <table width="100%" cellspacing="6" border="0" cellpadding="6">
<%
            if SecurityCheck(1) = true then ' User must have write access for this department
%> 
            <tr bgcolor="#eeeeee"> 
              <td bgcolor="#eeeeee">
                <font size="2" face="Arial" color="#3366cc"><b>Current Draft Plan:</b><br><font color="#333333" size="2" face="Arial">
<%
                if not rsDraft.BOF then
                
                  while not rsDraft.EOF
                    if rsDraft("apFormACompleted") = true then
%>
                    <a href="javascript:void(0)" onclick="javascript:OpenWindow('ActionPlanFormB.asp?apID=<%=rsDraft("apID")%>');"><%=rsDraft("apStartYear")%> - <%=rsDraft("apEndYear")%></a>
					<!-- displays a printer icon for printing a draft version of the ActionPlanFormB (EHS Plan) that shows all form field contents etc. CL 3/7/08 -->					&nbsp;&nbsp;<a href="javascript:void(0)" onclick="javascript:OpenWindow('ActionPlanReportDraft.asp?apID=<%=rsDraft("apID")%>');"  title="Click on the printer icon to view a print-friendly version of the draft Health and Safety Plan."><img src="printericon.gif" alt="Print-friendly format" width="16" height="16" border="0"></a>
					<!-- displays a printer icon for printing a draft version of the ActionPlanFormB (EHS Plan) that shows all form field contents etc. CL 3/7/08 -->
					<br>
<%
                    else 
%>
                    <a href="javascript:void(0)" onclick="javascript:OpenWindow('ActionPlanFormA.asp?apID=<%=rsDraft("apID")%>');"><%=rsDraft("apStartYear")%> - <%=rsDraft("apEndYear")%></a><br>
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
            End if
%>
            <tr bgcolor="#eeeeee"> 
              <td bgcolor="#eeeeee"> 
                <font size="2" face="Arial" color="#3366cc"><b>Final Plans:</b></font><br><font color="#333333" size="2" face="Arial">
<%
                if not rsFinal.BOF then
                  while not rsFinal.EOF
%>
                    <a href="javascript:void(0)" onclick="javascript:OpenWindow('ActionPlanReport.asp?apID=<%=rsFinal("apID")%>');"><%=rsFinal("apStartYear")%> - <%=rsFinal("apEndYear")%></a><br>
<%
                    rsFinal.movenext
                  wend
                end if
%>
                </font> 
              </td>
            </tr>
<%
            if SecurityCheck(3) = true then ' User must have write access for this department
%> 
<!-- removed since this is not working - should show service agreements       
			<tr bgcolor="#eeeeee"> 
              <td bgcolor="#eeeeee">
                <font size="2" face="Arial" color="#3366cc"><b>Current Draft Service Agreement:</b></font><br><font color="#333333" size="2" face="Arial">
<%
                if not rsSADraft.BOF then
                  while not rsSADraft.EOF
%>
                    <a href="javascript:void(0)" onclick="javascript:OpenWindow('ServiceAgreementForm.asp?apID=<%=rsSADraft("apID")%>&saID=<%=rsSADraft("saID")%>');"><%=rsSADraft("apStartYear")%></a>
					
					
<!-- displays a printer-friendly icon for the printing the draft Service Agreement in admin mode only. CL 3/7/2008
<a href="javascript:void(0)" onclick="javascript:OpenWindow('ServiceAgreementReportDraft.asp?apID=<%=rsSADraft("apID")%>&saID=<%=rsSADraft("saID")%>');" title="Click on the printer icon to view a print-friendly version of the Service Agreement."><img src="printericon.gif" alt="Print-friendly format" width="16" height="16" border="0"></a>					
<!-- end of the printer-friendly Service Agreement section. CL 3/7/2008

					<br>
<%
                    rsSADraft.movenext
                  wend
                end if
%>
                </font>
              </td>
            </tr>
 -->

<%
            else

              if SecurityCheck(1) = true then 
%>            
<!-- removed since this is not working - should show service agreements -->
			  <!-- tr bgcolor="#eeeeee">  
                <td bgcolor="#eeeeee"><b><font size="2" face="Arial" color="#3366cc">Current Draft Service Agreement :</font></b><br>
                <font color="#333333" size="2" face="Arial">
<%
                if not rsSADraft.BOF then
%>
                    <a href="javascript:void(0)" onclick="javascript:OpenWindow('ServiceAgreementReport.asp?apID=<%=rsSADraft("apID")%>&saID=<%=rsSADraft("saID")%>&draft=true');"><%=rsSADraft("apStartYear")%></a><br>
<%
                end if
%>
                </font>                   
                </td>
              </tr -->
<%
              end if
            end if
            
            if SecurityCheck(4) = true then 'all faculty/unit users can view final Service Agreement DLJ
%>          
<!-- removed since this is not working - should show service agreements
            <tr bgcolor="#eeeeee"> 
              <td bgcolor="#eeeeee"> 
                <font size="2" face="Arial" color="#3366cc"><b>Final Saved Service Agreement:</b></font>
								<br>
								<font color="#333333" size="2" face="Arial"> 
<%
                if not rsSAFinal.BOF then
                  while not rsSAFinal.EOF
%>
                    <a href="javascript:void(0)" onclick="javascript:OpenWindow('ServiceAgreementReport.asp?apID=<%=rsSAFinal("apID")%>&saID=<%=rsSAFinal("saID")%>');"><%=rsSAFinal("apStartYear")%></a>&nbsp;&nbsp;
<!-- added by CL/EHS 7/6/2007 
										<a href="javascript:void(0)" onclick="javascript:OpenWindow('ServiceAgreementStatus.asp?apID=<%=rsSAFinal("apID")%>&saID=<%=rsSAFinal("saID")%>');" title="Click on the printer icon to view a print-friendly version of the Service Agreement."><img src="printericon.gif" alt="Print-friendly format" width="16" height="16" border="0"></a><br>
<!-- end of the print-friendly service agreement code CL/EHS 7/6/2007 
<%
                    rsSAFinal.movenext
                  wend
                end if
%>
                </font>
              </td>
            </tr>
-->

<%                      
				Else
%>
            <!--tr bgcolor="#eeeeee"> 
              <td bgcolor="#eeeeee"> 
                <font size="2" face="Arial" color="#3366cc"><b>Final Saved Service Agreement:</b></font><br><font color="#333333" size="2" face="Arial"> 
<%
                if not rsSAFinal.BOF then
                  while not rsSAFinal.EOF
%>
                    <a href="javascript:void(0)" onclick="javascript:OpenWindow('ServiceAgreementReportNA.asp?apID=<%=rsSAFinal("apID")%>&saID=<%=rsSAFinal("saID")%>');"><%=rsSAFinal("apStartYear")%></a><br>
<%
                    rsSAFinal.movenext
                  wend
                end if
%>
                </font>
              </td>
            </tr -->
<%            End if
            
            if ActionPlan <> "" then
%>
            <!-- tr bgcolor="#eeeeee"> 
              <td bgcolor="#eeeeee">
                <p><font color="#333333" size="2" face="Arial"><a href="javascript:void(0)" onclick="javascript:OpenWindow('AuditWorksheet.asp?apID=<%=ActionPlan%>');">Audit Worksheet</a></font> 
              </td>
            </tr>
            <tr bgcolor="#eeeeee"> 
              <td bgcolor="#eeeeee">
                <p><font color="#333333" size="2" face="Arial"> <a href="javascript:void(0)" onclick="javascript:OpenWindow('ComplianceWorksheet.asp?apID=<%=ActionPlan%>');">Link to Compliance Assessment Worksheet</a></font> 
              </td>
            </tr -->
<%
            end if
%>
            </table>
          </td>
        </tr>
        </table>
        
      </td>
    </tr>
    </table>
    
  </td>
</tr>
</table>

<!-- #Include file="include\footer.asp" -->