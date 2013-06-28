<!--#include file="JSON_2.0.4.asp"-->
<!--#include file="JSON_UTIL_0.1.1.asp"-->
<%
Dim dbConn

Set dbConn = Server.CreateObject("ADODB.Connection")
dbConn.Open "Provider=SQLOLEDB;Initial Catalog=GISOnline;Data Source=10.158.34.35;User ID=gisonline;Password=gisonline;"

QueryToJSON(dbConn, "SELECT * FROM CIO_DIGCOM ORDER BY NAME").Flush

%>
