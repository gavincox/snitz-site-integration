<%
'#################################################################################
'## Copyright (C) 2000-02 Michael Anderson, Pierre Gorissen,
'##                       Huw Reddick and Richard Kinser
'##
'## This program is free software; you can redistribute it and/or
'## modify it under the terms of the GNU General Public License
'## as published by the Free Software Foundation; either version 2
'## of the License, or any later version.
'##
'## All copyright notices regarding Snitz Forums 2000
'## must remain intact in the scripts and in the outputted HTML
'## The "powered by" text/logo with a link back to
'## http://forum.snitz.com in the footer of the pages MUST
'## remain visible when the pages are viewed on the internet or intranet.
'##
'## This program is distributed in the hope that it will be useful,
'## but WITHOUT ANY WARRANTY; without even the implied warranty of
'## MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'## GNU General Public License for more details.
'##
'## You should have received a copy of the GNU General Public License
'## along with this program; if not, write to the Free Software
'## Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'##
'## Support can be obtained from support forums at:
'## http://forum.snitz.com
'##
'## Correspondence and Marketing Questions can be sent to:
'## reinhold@bigfoot.com
'##
'## or
'##
'## Snitz Communications
'## C/O: Michael Anderson
'## PO Box 200
'## Harpswell, ME 04079
'#################################################################################
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_admin.asp" -->
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
%>
<script language="javascript">

function doDefaultHighlights(){
	// Restore highlighted cells to form defaults.
	if (document.Form1.strSiteIntegEnabled[0].defaultChecked) 
		{
		if (document.Form1.strSiteLeft[0].defaultChecked)
			siteleft.style.backgroundColor='#99ff99';
		else
			siteleft.style.backgroundColor='#ff9999';

		if (document.Form1.strSiteHeader[0].defaultChecked)
			siteheader.style.backgroundColor = '#99ff99';
		else
			siteheader.style.backgroundColor = '#ff9999';
		
		if (document.Form1.strSiteRight[0].defaultChecked)
			siteright.style.backgroundColor = '#99ff99';
		else
			siteright.style.backgroundColor = '#ff9999';
		
		if (document.Form1.strSiteFooter[0].defaultChecked)
			sitefooter.style.backgroundColor = '#99ff99';
		else
			sitefooter.style.backgroundColor = '#ff9999';
			
		if (document.Form1.strSiteBorder[0].defaultChecked)
			disptable.border = '1';
		else
			disptable.border = '0';
			
		}
		else {
			siteleft.style.backgroundColor='#cccccc';
			siteheader.style.backgroundColor='#cccccc';
			sitefooter.style.backgroundColor='#cccccc';
			siteright.style.backgroundColor='#cccccc';
			disptable.border = '0';
		}
}

function doHighlights(){
	// Set highlighted cells to current checkbox values.

	if (document.Form1.strSiteIntegEnabled[0].checked) 
		{
		if (document.Form1.strSiteLeft[0].checked)
			siteleft.style.backgroundColor='#99ff99';
		else
			siteleft.style.backgroundColor='#ff9999';

		if (document.Form1.strSiteHeader[0].checked)
			siteheader.style.backgroundColor = '#99ff99';
		else
			siteheader.style.backgroundColor = '#ff9999';
		
		if (document.Form1.strSiteRight[0].checked)
			siteright.style.backgroundColor = '#99ff99';
		else
			siteright.style.backgroundColor = '#ff9999';
		
		if (document.Form1.strSiteFooter[0].checked)
			sitefooter.style.backgroundColor = '#99ff99';
		else
			sitefooter.style.backgroundColor = '#ff9999';
		
		if (document.Form1.strSiteBorder[0].checked)
			disptable.border = '1';
		else
			disptable.border = '0';
			
		}		
		else {
			siteleft.style.backgroundColor='#cccccc';
			siteheader.style.backgroundColor='#cccccc';
			sitefooter.style.backgroundColor='#cccccc';
			siteright.style.backgroundColor='#cccccc';
			disptable.border = '0'
		}
}
</script>
<%


Response.Write	"      <table border=""0"" width=""100%"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All&nbsp;Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Site&nbsp;Integration&nbsp;Configuration<br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine

if Request.Form("Method_Type") = "Write_Configuration" then
	Err_Msg = ""

	if Err_Msg = "" then
		for each key in Request.Form
			if left(key,3) = "str" or left(key,3) = "int" then
				strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLstring"))
			end if
		next

		Application(strCookieURL & "ConfigLoaded") = ""

		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Configuration Posted!</font></p>" & vbNewLine & _
		"      <meta http-equiv=""Refresh"" content=""2; URL=admin_home.asp"">" & vbNewLine & _
		"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Congratulations!</font></p>" & vbNewLine & _
		"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""admin_home.asp"">Back To Admin Home</font></a></p>" & vbNewLine
	else
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
		"      <table align=""center"" border=""0"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
	end if
else
	Response.Write	"    <form action=""admin_config_integration.asp"" method=""post"" id=""Form1"" name=""Form1"" onReset=""Javascript:doDefaultHighlights();"">" & vbNewLine & _
	"    <input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
	"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
	"        <tr>" & vbNewLine & _
	"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
	"            <table border=""0"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
	"              <tr valign=""middle"">" & vbNewLine & _
	"                <td background=""" & strImageURL & strHeadCellBGImage & """ bgcolor=""" & strHeadCellColor & """ colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><b>Site Integration Configuration</b></font></td>" & vbNewLine & _
	"              </tr>" & vbNewLine & _
	"              <tr valign=""middle"">" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Site Integration features:</b>&nbsp;</font></td>" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
	"				 On: <input type=""radio"" class=""radio"" name=""strSiteIntegEnabled"" value=""1""" & chkRadio(strSiteIntegEnabled,0,false) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
	"                Off: <input type=""radio"" class=""radio"" name=""strSiteIntegEnabled"" value=""0""" & chkRadio(strSiteIntegEnabled,0,true) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
	"                </font>&nbsp;</td>" & vbNewLine & _
	"              </tr>" & vbNewLine & _
	
	"              <tr valign=""middle"">" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Version info:</b>&nbsp;</font></td>" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
	"					[<em>v1.1</em>]" & vbNewLine & _
	"				 </font>&nbsp;</td>" & vbNewLine & _
	"              </tr>" & vbNewLine & _
	
	"              <tr valign=""middle"">" & vbNewLine & _
	"                <td width=""400"" bgColor=""" & strPopUpTableColor & """ align=""center"" colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" & vbNewLine & _
	"                <br /><table id=""disptable"" width=""380"" border="""
	if strSiteBorder = "1" then
		response.write "1"
	else
		response.write "0"
	end if
	response.write """ cellpadding=""5"" cellspacing=""1"">" & vbNewLine & _
	"                <tr>" & vbNewLine & _
	"                	<td name=""siteheader"" id=""siteheader"" colspan=""3"" align=""center"""
	if strSiteIntegEnabled = "0" then
		response.write " bgcolor=""#cccccc"""
	else
		if strSiteHeader = "1" then
			response.write " bgcolor=""#99ff99"""
		else
			response.write " bgcolor=""#ff9999"""
		end if
	end if
	response.write  "><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>Site Header</font></td></tr>" & vbNewLine & _
	"                <tr>" & vbNewLine & _
	"                	<td id=""siteleft"" width=""40"" align=""center"""
	if strSiteIntegEnabled = "0" then
		response.write " bgcolor=""#cccccc"""
	else
		if strSiteLeft = "1" then
			response.write " bgcolor=""#99ff99"""
		else
			response.write " bgcolor=""#ff9999"""
		end if
	end if
	response.write  "><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>Site<br />Left</font></td>" & vbNewLine & _
	"                	<td width=""300"" align=""center"" bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><img src=""" & strImageURL & "logo_snitz_forums_2000.gif"" alt=""Forums"" width=""163"" height=""76"" border=""0""></font></td>" & vbNewLine & _
	"                	<td id=""siteright"" width=""40"" align=""center"""
	if strSiteIntegEnabled = "0" then
		response.write " bgcolor=""#cccccc"""
	else
		if strSiteRight = "1" then
			response.write " bgcolor=""#99ff99"""
		else
			response.write " bgcolor=""#ff9999"""
		end if
	end if
	response.write  "><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>Site<br />Right</font></td></tr>" & vbNewLine & _
	"                <tr>" & vbNewLine & _
	"                	<td id=""sitefooter"" colspan=""3"" align=""center"""
	if strSiteIntegEnabled = "0" then
		response.write " bgcolor=""#cccccc"""
	else
		if strSiteFooter = "1" then
			response.write " bgcolor=""#99ff99"""
		else
			response.write " bgcolor=""#ff9999"""
		end if
	end if
	response.write "><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>Site Footer</font></td></tr>" & vbNewLine & _
	"                </table><br />" & vbNewLine & _
	"               </td>" & vbNewLine & _
	"              </tr>" & vbNewLine & _
	
	"              <tr valign=""middle"">" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Site Header:</b>&nbsp;</font></td>" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
	"				 On: <input type=""radio"" class=""radio"" name=""strSiteHeader"" value=""1""" & chkRadio(strSiteHeader,0,false) & " onClick=""javascript:doHighlights();"">" & vbNewLine & _
	"                Off: <input type=""radio"" class=""radio"" name=""strSiteHeader"" value=""0""" & chkRadio(strSiteHeader,0,true) & " onClick=""javascript:doHighlights();"">" & vbNewLine & _
	"                </font>&nbsp;</td>" & vbNewLine & _
	"              </tr>" & vbNewLine & _
	
	"              <tr valign=""middle"">" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Site Left:</b>&nbsp;</font></td>" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
	"				 On: <input type=""radio"" class=""radio"" name=""strSiteLeft"" value=""1""" & chkRadio(strSiteLeft,0,false) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
	"                Off: <input type=""radio"" class=""radio"" name=""strSiteLeft"" value=""0""" & chkRadio(strSiteLeft,0,true) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
	"                </font>&nbsp;</td>" & vbNewLine & _
	"              </tr>" & vbNewLine & _
	
	"              <tr valign=""middle"">" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Site Right:</b>&nbsp;</font></td>" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
	"				 On: <input type=""radio"" class=""radio"" name=""strSiteRight"" value=""1""" & chkRadio(strSiteRight,0,false) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
	"                Off: <input type=""radio"" class=""radio"" name=""strSiteRight"" value=""0""" & chkRadio(strSiteRight,0,true) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
	"                </font>&nbsp;</td>" & vbNewLine & _
	"              </tr>" & vbNewLine & _
	
	"              <tr valign=""middle"">" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Site Footer:</b>&nbsp;</font></td>" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
	"				 On: <input type=""radio"" class=""radio"" name=""strSiteFooter"" value=""1""" & chkRadio(strSiteFooter,0,false) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
	"                Off: <input type=""radio"" class=""radio"" name=""strSiteFooter"" value=""0""" & chkRadio(strSiteFooter,0,true) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
	"                </font>&nbsp;</td>" & vbNewLine & _
	"              </tr>" & vbNewLine & _
	
	"              <tr valign=""middle"">" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Show Borders:</b>&nbsp;</font></td>" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
	"				 On: <input type=""radio"" class=""radio"" name=""strSiteBorder"" value=""1""" & chkRadio(strSiteBorder,0,false) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
	"                Off: <input type=""radio"" class=""radio"" name=""strSiteBorder"" value=""0""" & chkRadio(strSiteBorder,0,true) & " onClick=""Javascript:doHighlights();"">" & vbNewLine & _
	"                </font>&nbsp;</td>" & vbNewLine & _
	"              </tr>" & vbNewLine & _
	
	"              <tr valign=""middle"">" & vbNewLine & _
	"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><input type=""submit"" value=""Submit New Config"" id=""submit1"" name=""submit1""> <input type=""reset"" value=""Reset Old Values"" id=""reset1"" name=""reset1""></td>" & vbNewLine & _
	"              </tr>" & vbNewLine & _
	"            </table>" & vbNewLine & _
	"          </td>" & vbNewLine & _
	"        </tr>" & vbNewLine & _
	"      </table>" & vbNewLine & _
	"    </form>" & vbNewLine
end if
WriteFooter
Response.End
%>
