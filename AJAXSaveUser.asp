<!--#include file="adovbs.inc"--> 
<%

	dim id, name, password, mode, department
	mode = request("mode")
	id = request("id")
	name = request("name")
	password = request("password")
	department = request("department")

	set con	= server.createobject ("adodb.connection")
		con.open "DSN=ehs"
		
	Set objCmd  = Server.CreateObject("ADODB.Command")
	objCmd.CommandType = adCmdText
	Set objCmd.ActiveConnection = con
		  
		  
	if mode = "edit" then
		 ' Get the ID of the ActionPlan that the FA and CC worksheets will be attached to
		  sqlUpdate = "update AD_Users set lgName = ?, lgPassword = ?, lgDepartment = ? where lgID = ?"  
		  objCmd.CommandText = sqlUpdate
		 
		  objCmd.Parameters.Append objCmd.CreateParameter("lgName", adWChar, adParamInput, 50)
		  objCmd.Parameters("lgName") = name
		  objCmd.Parameters.Append objCmd.CreateParameter("lgPassword", adWChar, adParamInput, 50)
		  objCmd.Parameters("lgPassword") = password
		  
		  objCmd.Parameters.Append objCmd.CreateParameter("lgDepartment", adInteger, adParamInput, 50)
		  objCmd.Parameters("lgDepartment") = cint(department)
		   objCmd.Parameters.Append objCmd.CreateParameter("lgID", adInteger, adParamInput, 50)
		  objCmd.Parameters("lgID") = cint(id)
		  
		  set rsFinal =  server.createobject("adodb.recordset")
		  rsFinal.Open objCmd
		  
		  Response.write(1)
	end if
	
	if mode = "new" then
		sqlNew = "insert into AD_Users (lgName, lgPassword, lgDepartment, lgView, lgEdit, lgChangePassword, lgSuperUser, lgAuditor) values (?,?,?,true,true,true,true,true)"
		objCmd.CommandText = sqlNew
		 
		objCmd.Parameters.Append objCmd.CreateParameter("lgName", adWChar, adParamInput, 50)
		objCmd.Parameters("lgName") = name
		objCmd.Parameters.Append objCmd.CreateParameter("lgPassword", adWChar, adParamInput, 50)
		objCmd.Parameters("lgPassword") = password
		 objCmd.Parameters.Append objCmd.CreateParameter("lgDepartment", adInteger, adParamInput, 50)
		  objCmd.Parameters("lgDepartment") = cint(department)
		

		set rsFinal =  server.createobject("adodb.recordset")
		rsFinal.Open objCmd

		Response.write(1)
	end if
	
	if mode = "delete" then
		sqlDel = "delete from AD_Users where lgID = ?"
		
		objCmd.CommandText = sqlDel
		 
		objCmd.Parameters.Append objCmd.CreateParameter("lgID", adInteger, adParamInput, 50)
		objCmd.Parameters("lgID") = cint(id)


		set rsFinal =  server.createobject("adodb.recordset")
		rsFinal.Open objCmd

		Response.write(1)
	end if
	
%>