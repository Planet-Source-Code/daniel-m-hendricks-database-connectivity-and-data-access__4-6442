<div align="center">

## Database Connectivity and Data Access


</div>

### Description

This article shows various connection strings, used to connect to various databases in Windows, as well as methods to access and modify data. Some connection strings may require client software to be installed, but most work with Windows 2000.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel M\. Hendricks](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-m-hendricks.md)
**Level**          |Beginner
**User Rating**    |4.5 (85 globes from 19 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-m-hendricks-database-connectivity-and-data-access__4-6442/archive/master.zip)





### Source Code

<b><font size="6">Database Connectivity in ASP</font></b>
<p><font size="3">This reference will show you how to connect to a variety of
databases in different ways:</font></p>
<ol>
 <li><a href="#connect">Connect to the Database</a>
 <li><a href="#run">Run your SQL commands</a>
 <li><a href="#examples">Common Examples</a></li>
</ol>
<p><b><font size="4"><u><a name="connect">Connect to the Database</a></u></font></b></p>
<p><font size="3">Before you can access your database, you need to connect to it
using one of the following methods: </font></p>
<p><font size="3"><b>Microsoft Access 2000 Database (OLE-DB):</b></font></p>
<p><font face="Courier New" size="2">Set db =
Server.CreateObject("ADODB.Connection")<br>
db.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" &
Server.MapPath("database.mdb") & ";"</font></p>
<p><i><font size="2">Microsoft Access databases are quick, easy, and portable.
It works good for small, intradepartmental applications. If you plan on
having more than a few users connecting to it, however, you many wish to
consider using a database like SQL Server or Oracle instead. Here is
another way to connect to a Microsoft Access database:</font></i></p>
<p><font size="3"><b>Microsoft Access 2000 Database:</b></font></p>
<p><font face="Courier New" size="2">Set db =
Server.CreateObject("ADODB.Connection")<br>
db.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" &
Server.MapPath("database.mdb") & ";UID=;PWD="</font></p>
<p><i><font size="2">For a description of the difference between OLE-DB and
ODBC, check out <a href="http://www.oledb.com/ole-db/index.html">this article</a>
at oledb.com. </font></i></p>
<p><b>Connecting to a database using a DSN:</b></p>
<p><font face="Courier New" size="2">Set db =
Server.CreateObject("ADODB.Connection")<br>
db.Open "DSN=mydsn;UID=username;PWD=password"</font></p>
<p><font size="2"><i>Before you can use this method, you must create a </i><b><i>DSN</i></b><i>
in your control panel (usually under ODBC or Data Sources). This process
varies from each version of Windows, so you're on your own. When you
create a DSN, you will be asked to give it a name. The name you enter
should replace the "mydsn" value above, along with the username and
password. </i></font></p>
<p><font size="3"><b>Connect to a SQL Server database with OLE DB:</b> </font></p>
<p><font face="Courier New" size="2">Set db =
Server.CreateObject("ADODB.Connection")<br>
db.Open "Provider=SQLOLEDB; Data Source=SERVER; Initial Catalog=database;
User ID=username; Password=password"</font></p>
<p><i><font size="2">An OLE DB connection can provide faster performance than a
DSN. This method doesn't require you to set up a DSN (which makes reloading the
machine easier), which makes it easier to reload the computer and doesn't
require you to create a DSN. However, if you move your applications to
another server or if you move your database to another server, you will need to
update any hard-coded values. There are ways around this, but for
simplicity, I have provided the example above. </font></i></p>
<p><font size="3"><b>Connect to a MySQL Database Under Linux/Chili!Soft ASP:</b> </font></p>
<p><font face="Courier New" size="2">Set db =
Server.CreateObject("ADODB.Connection")<br>
db.Open "Driver={MySQL}; SERVER=localhost; DATABASE=database; UID=username;
PWD=password"</font></p>
<p><i><font size="2">This code has only been tested on a Cobalt RAQ with
Chili!Soft ASP and MySQL.</font></i></p>
<p><font size="3"><b>Connect to Oracle 8 (OLE-DB):</b> </font></p>
<p><font face="Courier New" size="2">Set db =
Server.CreateObject("ADODB.Connection")<br>
db.Open "Provider=OraOLEDB.Oracle;User ID=user;Password=pwd; Data Source=hoststring;"</font></p>
<p><i><font size="2">This code has only been confirmed to work with Oracle 8i
server and Windows client. Important: Requires Oracle client connectivity
tools to be installed. Here is another way to connect to an Oracle
database:</font></i></p>
<p><font size="3"><b>Connect to Oracle 8:</b> </font></p>
<p><font face="Courier New" size="2">Set db =
Server.CreateObject("ADODB.Connection")<br>
db.Open "Driver={Microsoft ODBC for Oracle};UID=user;PWD=password;CONNECTSTRING=hoststring"</font></p>
<p><i><font size="2">This also requires the Oracle client tools be
installed. For a description of the difference between OLE-DB and ODBC,
check out <a href="http://www.oledb.com/ole-db/index.html">this article</a> at
oledb.com. </font></i></p>
<p><b><font size="4"><u><br>
<a name="run">Run Your Commands</a></u></font></b></p>
<p><font size="3">Now that you have a connection to your database, you can run
SQL statements:</font></p>
<p><b><font size="3">Delete Records:</font></b></p>
<p><font face="Courier New" size="2">db.execute("DELETE FROM mytable WHERE
FullName = 'John Doe'")</font></p>
<p><i><font size="2">This is only used as an example. You will need to
replace "mytable" with the name of the table you are trying to delete
from. Likewise, replace "FullName" with the name of the
appropriate field.</font></i></p>
<p><b><font size="3">Insert Records:</font></b></p>
<p><font face="Courier New" size="2">db.execute("INSERT INTO mytable VALUES
('John Doe', 22, '321 Disk Dr.', 'Hollywood, CA')</font></p>
<p><font size="2"><i>Again, this is only used as an example. Change the
statement as needed.</i></font></p>
<p><b><font size="3">List Records:</font></b></p>
<p><font face="Courier New" size="2">set rs=db.execute("SELECT * FROM
mytable")<br>
rs.MoveFirst<br>
Do Until rs.EOF<br>
   Response.Write rs("MyField") & &quot;&lt;br&gt;&quot;<br>
Loop</font></p>
<p><i><font size="2">The first line is a select statement that selects records.
The following lines iterate through each line, displays the current value of the
"MyField" field, and adds a line-feed. You will want to change
the "mytable" and "MyField" values appropriately. </font></i></p>
<p><b><font size="4"><u><br>
<a name="examples">Common Examples</a></u></font></b></p>
<p><b><font size="3">Add, list, and delete records:</font></b><br>
<table border="0">
 <tbody>
 <tr>
  <td width="50"></td>
  <td><font face="Courier New" size="2"><br>
  <br>
  Set db = Server.CreateObject("ADODB.Connection")<br>
  db.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" &
  Server.MapPath("database.mdb") & ";"<br>
  <br>
  db.execute("INSERT INTO MyTable VALUES ('Dan Hendricks', 22)")<br>
  set rs=db.execute("SELECT * FROM MyTable")<br>
  <br>
  rs.MoveFirst<br>
  Do Until rs.EOF<br>
  Response.Write rs("NAME") & &quot;&lt;br&gt;&quot;<br>
  rs.MoveNext<br>
  Loop<br>
  <br>
  db.execute("DELETE FROM MyTable WHERE NAME = 'Dan Hendricks'")<br>
  </font></td>
 </tr>
 </tbody>
</table>
<p><font size="2"><i>This code will open the database, add the values "Dan
Hendricks" and "22" into the first two field of the chosen table,
display all current records in the table, and finally delete the record that was
added.</i></font></p>
<p><b><font size="3">Here is another quick and easy way to connect and list
records:</font></b></p>
<table border="0">
 <tbody>
 <tr>
  <td width="50"></td>
  <td>
  <p><font face="Courier New" size="2">'This code connects to the
  database.<br>
  set rs=Server.CreateObject("ADODB.Recordset")<br>
  db="DSN=TechSupport;UID=TechSupport;PWD=foobar"</font></p>
  <p><font face="Courier New" size="2">'This code iterates through the
  current records.<br>
  mySQL = "SELECT * from chairs "<br>
  rs.open mySQL, db, 1, 3  <!-- Change the '3' to a '1' for
  a read-only. --><br>
  rs.MoveFirst<br>
  Do Until rs.EOF<br>
     Response.Write rs("MyField") &
  &quot;&lt;br&gt;&quot;<br>
     rs.MoveNext<br>
  Loop</font></p>
  <p><font face="Courier New" size="2">'This code deletes a record, and
  then adds a new one<br>
  rs.MoveFirst<br>
  rs.Delete<br>
  rs.AddNew<br>
    rs("Name") = 'Jane Doe'<br>
  rs.Update<br>
  rs.Close</font></p>
  </td>
 </tr>
 </tbody>
</table>
<p><i><font size="2">NOTE: This does not use the same connect statements
listed above. It's just a different way to connect to a database and list,
add, or remove records.</font></i>

