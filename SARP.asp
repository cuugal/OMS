<%@language = VBscript%>
<!-- #Include file="include\general.asp" -->
<!--#INCLUDE FILE="DbConfig.asp"-->

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

	'set con	= server.createobject ("adodb.connection")
	'con.open "DSN=ehs"
	
	   
  set conn = Server.CreateObject("ADODB.Connection")
  conn.open constr
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
   
   temp = instr(1,strComments,"'",vbTextCompare)
      if temp > 1 then 
         strComments = Replace(strComments,"'","''",1)
         Response.Write(strComments)%><BR><%
         Response.Write("Value")
         Response.Write(temp) 
      end if
   
  ' response.write(numReqId)%><%
  'response.write(strServiceAction)
  
   if strServiceAction = "on" then %><%
          strServiceAction = "Yes"
          'response.write(strServiceAction)
          'response.write(numReqId)
          %><%

          else
          strServiceAction = "No"
      %><%
      
     
    end if
   strSQL = "UPDATE SA_ServiceAgreementDetails SET sdServiceActioned = "&strServiceAction&", sdComments = '"& strComments &"' where sdServiceAgreement = "&servAgr&" and sdRequirement ="&numReqId 
   
   'Response.Write(strSQL)
   
   'set rsCheckCampus = Server.CreateObject("ADODB.Recordset")
                      'rsCheckCampus.Open strSQL, conn, 3, 3

'************************************loop ends here******************************************************	
set rsAdd =  server.CreateObject("ADODB.Recordset")
 rsAdd.Open strSQL,conn  ' need to look at this

next
      if SA = "on" then %><%
          sA = "Yes"
          else
          SA = "No"
      end if
   strSQL1 = "UPDATE SA_ServiceAgreement SET saSA = "&SA&",saC ='"&C&"' where saActionPlan = "&AP
  set rsAddSA = server.CreateObject("ADODB.Recordset")
  rsADDSA.Open strSQL1,conn,3,3
  
Response.Write ("The Service Agreement has been Updated")
'Response.end 

'closeWindow()
end if 
'********************************************************************************************************%>
<!-- #Include file="include\footer.asp" -->