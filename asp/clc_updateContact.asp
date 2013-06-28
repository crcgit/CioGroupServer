<%
Dim DataConn,SQL,RS,imgstr,id,bio,name,email,citycounty,phoneoffice,phonemobile,phonehome,password,result

imgstr = Request.QueryString("im")
id = Request.QueryString("id")
bio = REPLACE(Request.QueryString("bi"), "'", "''")
name = REPLACE(Request.QueryString("na"), "'", "''")
email = REPLACE(Request.QueryString("em"), "'", "''")
citycounty = REPLACE(Request.QueryString("cc"), "'", "''")
phoneoffice = REPLACE(Request.QueryString("po"), "'", "''")
phonemobile = REPLACE(Request.QueryString("pm"), "'", "''")
phonehome = REPLACE(Request.QueryString("ph"), "'", "''")
password = REPLACE(Request.QueryString("pw"), "'", "''")


Set DataConn = Server.CreateObject("ADODB.Connection")

ConnStr = "Provider=SQLOLEDB;Initial Catalog=GISOnline;Data Source=10.158.34.35;User ID=gisonline;Password=gisonline;"

SQL = "UPDATE CIO_DIGCOM SET IMG_STRING = '" & imgstr 
SQL = SQL + "', BIOGRAPHY = '" & bio 
SQL = SQL + "', NAME = '" & name 
SQL = SQL + "', EMAIL = '" & email 
SQL = SQL + "', CITY_COUNTY = '" & citycounty 
SQL = SQL + "', PHONE_OFFICE = '" & phoneoffice 
SQL = SQL + "', PHONE_MOBILE = '" & phonemobile 
SQL = SQL + "', PHONE_HOME = '" & phonehome 
SQL = SQL + "', PASSWORD = '" & password 
SQL = SQL + "' WHERE ID = '" & id & "'"

DataConn.Open ConnStr
DataConn.execute SQL

DataConn.Close
Set DataConn = Nothing

result = "OK"
response.write(result)
	   %>
