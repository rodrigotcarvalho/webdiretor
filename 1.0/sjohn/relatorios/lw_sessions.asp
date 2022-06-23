<%


IF Request.QueryString("setVar") = "true" Then
	Session("LOCAWEB_Data") = Now()
End IF

For Each item in Session.Contents

	Response.Write(item & ": " & Session(item) & "<br/>")

Next

%>