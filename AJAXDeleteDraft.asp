<!--#include file="adovbs.inc"--> 
<%


	apid = request("apid")


	set con	= server.createobject ("adodb.connection")
		con.open "DSN=ehs"
		
	Set objCmd  = Server.CreateObject("ADODB.Command")
	objCmd.CommandType = adCmdText
	Set objCmd.ActiveConnection = con
		  

	sqlDel = "delete from AP_ActionPlans where apid = ?"
		
	objCmd.CommandText = sqlDel
		 
	objCmd.Parameters.Append objCmd.CreateParameter("apid", adInteger, adParamInput, 50)
	objCmd.Parameters("apid") = cint(apid)


	set rsFinal =  server.createobject("adodb.recordset")
	rsFinal.Open objCmd

	Response.write(1)

	
%>