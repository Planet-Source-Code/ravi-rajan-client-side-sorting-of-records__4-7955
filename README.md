<div align="center">

## Client Side Sorting of records


</div>

### Description

THIS CODE WILL CONVERT A RECORDSET INTO A CLIENT SIDE ARRAY AND THIS ARRAY CAN BE SORTED AS PER ANY FIELD IN THE TABLE.SINCE THE PROCESSING IS DONE ON THE CLIENT SIDE ITSELF PERFORMANCE IS MUCH BETTER.CLICKING ON THE HYPERLINKED TABLE COLUMN HEADERS WILL CSORT THE RECORDS AS PER THAT COLUMN.
 
### More Info
 
JUST REPLACE THE CONNECTION STRING WITH YOUR SQL-SERVER CONNECTION AND THE CODE IS READY TO RUN

SORTED CLIENT SIDE ARRAY


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ravi Rajan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ravi-rajan.md)
**Level**          |Advanced
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ravi-rajan-client-side-sorting-of-records__4-7955/archive/master.zip)





### Source Code

```
<%
 dim conn
 dim strconn
 dim objrs
 dim strsql
 set conn=server.CreateObject("ADODB.connection")
 set objrs=server.CreateObject("ADODB.recordset")
 strconn="Provider=SQLOLEDB.1;Password=efc;Persist Security Info=True;User ID=test;Initial Catalog=Northwind;Data Source=SMART\EPRD"
 conn.ConnectionString=strconn
 conn.Open
strsql="SELECT P.PRODUCTNAME AS PRODUCT,C.CATEGORYNAME AS CATEGORY,S.COMPANYNAME AS SUPPLIER " & _
"FROM PRODUCTS P,CATEGORIES C,SUPPLIERS S WHERE " & _
"P.CATEGORYID=C.CATEGORYID AND " & _
"P.SUPPLIERID=S.SUPPLIERID " & _
"ORDER BY P.PRODUCTNAME "
objrs.Open strsql,conn
%>
<HTML>
<HEAD>
<SCRIPT LANGUAGE=javascript>
function product(product,category,supplier)
{
this.product=product;
this.category=category;
this.supplier=supplier;
}
</SCRIPT>
<%
'converting recordset into a client side array
dim nor
while not objrs.EOF
	nor=nor+1
	objrs.MoveNext
wend
	objrs.MoveFirst
	Response.Write("<script language=JavaScript>")
	Response.Write("var nor=" & nor &";")
	Response.Write("var arr_prod = new Array(" & nor &");")
	Response.Write("var index = 0;")
dim index
index=0
while not objrs.EOF
	Response.Write("arr_prod[" & index & "] = new product("""& objrs.Fields(0) & ""","""& objrs.Fields(1) & ""","""& objrs.Fields(2) & """);")
	index=index+1
	objrs.MoveNext
wend
	Response.Write("</script>")
%>
<style type="TEXT/CSS"><!--
h1, td, th { font-family: Arial; }
tr.color0 { background: #ccffcc; }
tr.color1 { background: #ccccff; }
tr.color2 { background: #ffcccc; }
//--></style>
</HEAD>
<BODY bgcolor="cyan">
<center><h2><u>Client Side Sorting Using <FONT COLOR=NAVY>ARRAY</FONT> Object's <FONT COLOR=NAVY>SORT</FONT> METHOD</u></h2></center>
	<table border="0" cellspacing="0" cellpadding="5" width="100%">
<tr class="color2">
<th><a href="ClientSideSorting.asp?product">Product</a></th>
<th><a href="ClientSideSorting.asp?category">Category</a></th>
<th><a href="ClientSideSorting.asp?supplier">Supplier</a></th>
</tr>
<script language="JavaScript"><!--
var output = '';
var searchtext = location.search.substring(1);
if (searchtext == '')
 product.prototype.toString = new Function('return this.product');
else
 product.prototype.toString =
  new Function('return this.' + searchtext + ';');
arr_prod.sort();
for (var i=0; i < arr_prod.length; i++) {
 output += '<tr class="color' + i%3 + '">';
 output += '<td align=center>' + arr_prod[i].product + '&nbsp;<\/td>';
 output += '<td align=center>' + arr_prod[i].category + '&nbsp;<\/td>';
 output += '<td align=center>' + arr_prod[i].supplier + '&nbsp;<\/td>';
 output += '<\/tr>';
}
document.write(output);
//--></script>
</table>
</BODY>
</HTML>
```

