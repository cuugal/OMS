<%
Dim constr
constr = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("Database/ehs.mdb")

Function InjectionEncode(str)
	InjectionEncode=Replace(str,"'","''")
End Function

%>