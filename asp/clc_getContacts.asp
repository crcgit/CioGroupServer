<!--#include file="JSON_2.0.4.asp"-->
<!--#include file="JSON_UTIL_0.1.1.asp"-->
<%
Dim dbConn

Set dbConn = Server.CreateObject("ADODB.Connection")
dbConn.Open "Provider=SQLOLEDB;Initial Catalog=CyberReliant_Guest;Data Source=54.225.81.158;User ID=guest;"

QueryToJSON(dbConn, "SELECT * FROM CIO_DIGCOM ORDER BY NAME").Flush

%>
