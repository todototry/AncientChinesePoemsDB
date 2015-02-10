<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>古今字音对照查询结果</title>
</head>

<body>
<%
dim connabc,objstr,objrs,sql,pinyin,hanzi,fanqie,she,denghu,sisheng,yunmu,shengmu
pinyin=request("pinyin")
hanzi=request("hanzi")
fanqie=request("fanqie")
she=request("she")
denghu=request("denghu")
sisheng=request("sisheng")
yunmu=request("yunmu")
shengmu=request("shengmu")
set connabc=server.CreateObject("ADODB.connection")
connabc.open "provider=Microsoft.jet.oledb.4.0;"&"data source="& server.MapPath("yin.mdb")
set objrs=server.createobject("ADODB.recordset")
sql="select * from 音表 where 拼音 like '%"& pinyin &"%' and 汉字 like '%"& hanzi &"%' and 反切 like '%"& fanqie &"%' and 摄 like '%"& she &"%' and 等呼 like '%"& denghu &"%' and 四声 like '%"& sisheng &"%' and 韵目 like '%"& yunmu &"%' and 声母 like '%"&shengmu &"%'"
objrs.open sql,connabc,1,3
%>
<%
if objrs.EOF then
   response.write "<p align=center><font color=red>没有查到相应的数据</font></p>"
else
%>
<table border=3 bordercolor=blue align=center><tr>
<%
for i=0 to objrs.fields.count-1
response.write"<th>"&objrs.fields(i).name&"</th>"
next
%>
</tr>
<%
do while not objrs.eof
data="<tr>"
for i=0 to objrs.fields.count-1
data=data&"<td>"&objrs.fields(i).value&"</td>"
next
response.write data&"</tr>"
objrs.MoveNext
   Loop
objrs.close 
Set objrs = Nothing
connabc.close 
set connabc=Nothing
end if
%>
</table>
<p align=center><a href=cha.htm>重新查询</a></p>
<p align=center><a href=cha2.htm>高级查询</a></p></body>
</html>
