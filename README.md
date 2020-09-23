<div align="center">

## Checking for installed server components


</div>

### Description

This code allows developers to know which components are installed on the server, based on list of 68 most common components.
 
### More Info
 
easy to extend...just add new string components on array and re-define array size, then run tha page again.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[FabricioC](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/fabricioc.md)
**Level**          |Beginner
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[ASP Server Object Model](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/asp-server-object-model__4-32.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/fabricioc-checking-for-installed-server-components__4-8976/archive/master.zip)





### Source Code

```
<%
dim aComponents(68)
aComponents(1) = "ADODB.Command"
aComponents(2) = "ADODB.Connection"
aComponents(3) = "ADODB.Recordset"
aComponents(4) = "ADODB.Stream"
aComponents(5) = "ADOX.Catalog"
aComponents(6) = "AspDNS.Lookup"
aComponents(7) = "ASPExec.Execute"
aComponents(8) = "AspHTTP.Conn"
aComponents(9) = "AspImage.Image"
aComponents(10) = "AspMX.Lookup"
aComponents(11) = "AspNNTP.Conn"
aComponents(12) = "AspPing.Conn"
aComponents(13) = "AspSock.Conn"
aComponents(14) = "CDO.MESSAGE"
aComponents(15) = "CDONTS.NewMail"
aComponents(16) = "Dundas.Mailer"
aComponents(17) = "Dundas.PieChartServer"
aComponents(18) = "Dundas.PieChartServer.2"
aComponents(19) = "Dundas.Upload"
aComponents(20) = "Dundas.Upload.2"
aComponents(21) = "Dundas.UploadProgress"
aComponents(22) = "ECHOCom.Echo"
aComponents(23) = "GuidMakr.GUID"
aComponents(24) = "ImgSize.Check"
aComponents(25) = "ixsso.Query"
aComponents(26) = "ixsso.Util"
aComponents(27) = "JMAil.Message"
aComponents(28) = "JMail.POP3"
aComponents(29) = "JMail.SMTPMail"
aComponents(30) = "JRO.JetEngine"
aComponents(31) = "Microsoft.DiskQuota.1"
aComponents(32) = "microsoft.XMLDOM"
aComponents(33) = "Microsoft.XMLHTTP"
aComponents(34) = "MSWC.AdRotator"
aComponents(35) = "MSWC.BrowserType"
aComponents(36) = "MSWC.ContentRotator"
aComponents(37) = "MSWC.Counters"
aComponents(38) = "MSWC.IISLog"
aComponents(39) = "MSWC.MyInfo"
aComponents(40) = "MSWC.MyInfo"
aComponents(41) = "MSWC.NextLink"
aComponents(42) = "MSWC.PageCounter"
aComponents(43) = "MSWC.PermissionChecker"
aComponents(44) = "MSWC.Status"
aComponents(45) = "MSWC.Tools"
aComponents(46) = "MSXML.DomDocument"
aComponents(47) = "MSXML2.DOMDocument"
aComponents(48) = "MSXML2.DOMDocument.3.0"
aComponents(49) = "Msxml2.FreeThreadedDOMDocument.3.0"
aComponents(50) = "MSXML2.ServerXMLHTTP"
aComponents(51) = "MSXML2.ServerXMLHTTP.3.0"
aComponents(52) = "MSXML2.XSLTemplate"
aComponents(53) = "Persits.Grid"
aComponents(54) = "Persits.Jpeg"
aComponents(55) = "Persits.MailSender"
aComponents(56) = "Persits.Upload"
aComponents(57) = "Persits.Upload.1"
aComponents(58) = "Persits.UploadProgress"
aComponents(59) = "POP3svg.Mailer"
aComponents(60) = "Scripting.Dictionary"
aComponents(61) = "Scripting.FileSystemObject"
aComponents(62) = "Scriptlet.TypeLib"
aComponents(63) = "SMTPsvg.Mailer"
aComponents(64) = "SOFTWING.AspTear"
aComponents(65) = "VBScript.RegExp"
aComponents(66) = "WinHttp.WinHttpRequest.5.1"
aComponents(67) = "WScript.Network"
aComponents(68) = "WScript.Shell"
Response.write("Installed components:<br><br>")
On error resume next
for i=1 to Ubound(aComponents)
 set obj = Server.CreateObject(aComponents(i))
 if err.number = 0 then
  Set obj = nothing
  Response.write(aComponents(i) & "<br>")
 end if
 err.clear
next
%>
```

