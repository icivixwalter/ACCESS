'------------------------------------------------------
' Setup the right pane with the 'default' source file -
'------------------------------------------------------
Sub SetupPage(myurl)
	iIsIE=SniffIt()   
	current=0
	 If iIsIE="4.0>" Then 
		  Document.All("gallery").Item(current).Style.Fontweight="Bold"
		 Document.All("gallery").Item(current).ParentElement.ParentElement.Children(0).InnerHTML="<img SRC='blkarrow.gif' WIDTH=4 HEIGHT=7>"
		  Document.All("gallery").Item(current).ParentElement.bgcolor="#FFCC00"
	 End If
	Top.Frames(1).Frames(2).Location.Href=myurl
End Sub  

'-------------------------------------------------------------------------
' The following 2 procedures handle visual effects on MouseOver/MouseOut -
' for the Gallery (G) type card                              -
'-------------------------------------------------------------------------
Sub G_Riser()
	If iIsIE="4.0>" Then 
		If cstr(Window.Event.srcElement.tagname)="TD" then
			Window.Event.srcElement.bgcolor="#FFFFCC"
		Else
			Window.Event.srcElement.ParentElement.bgcolor="#FFFFCC"
		End If
	End If
End Sub

Sub G_Flatten()
	If iIsIE="4.0>" Then 
		If Cstr(Window.Event.srcElement.tagname)="TD" Then
			 Window.Event.srcElement.bgcolor="#FFFFFF"
		Else
			 Window.Event.srcElement.ParentElement.bgcolor="#FFFFFF"
		End If
	 
	  Document.All("gallery").Item(current).Style.Fontweight="bold"
		Document.All("gallery").Item(current).ParentElement.bgcolor="#FFCC00"
		Document.All("gallery").Item(current).ParentElement.ParentElement.Children(0).InnerHTML="<IMG SRC='blkarrow.gif'>"
	End If
End Sub

'-----------------------------------------------------------------
' Load the content pane (right) with the appropriate source file -
'-----------------------------------------------------------------
Sub G_LoadPane(myurl)
	If iIsIE="4.0>" Then
		Window.Event.ReturnValue=False
		Document.All("gallery").Item(current).ParentElement.ParentElement.Children(0).InnerHTML=""
		Document.All("gallery").Item(current).Style.Fontweight="normal"
		Document.All("gallery").Item(current).ParentElement.bgcolor="#FFFFFF"

		 For q = 0 To Document.All("gallery").Length - 1
			  If Window.Event.srcElement.Contains(Document.All("gallery").Item(q)) Then
					  current=q
			  End If
		 Next

		 Document.All("gallery").Item(current).Style.Fontweight="bold"
		 Document.All("gallery").Item(current).ParentElement.ParentElement.Children(0).InnerHTML="<IMG SRC='blkarrow.gif'>"
		 Document.All("gallery").Item(current).ParentElement.bgcolor="#FFCC00"
	End If
	Top.Frames(4).Location.Href=myurl
End Sub

'------------------------
' Detect if IE is >=4.0 -
'------------------------
Function DetectBrowserVersion()
	Dim iVersion
	iVersion=navigator.appversion
	If Left(iVersion,1)>=4 Then
		 DetectBrowserVersion="4.0>"
	Else
		 DetectBrowserVersion="3.0x"
	End if
	
End Function

'-------------------------------------------------------------------------
' The following 2 procedures handle visual effects on MouseOver/MouseOut -
' for the Serie (S) type card                                -
'-------------------------------------------------------------------------
Sub S_Riser()
	If cstr(Window.Event.srcElement.tagname)="TD" then
		Window.Event.srcElement.bgcolor="#ffffcc"
	Else
		Window.Event.srcElement.parentElement.bgcolor="#ffffcc"
	End If
End Sub


Sub S_Flatten()
	If cstr(Window.Event.srcElement.tagname)="TD" then
		Window.Event.srcElement.bgcolor="#ffcc00"
	Else
		Window.Event.srcElement.parentElement.bgcolor="#ffcc00"
	End If
	Document.All("hotnumber").Item(current).Style.fontweight="bold"
	Document.All("hotnumber").Item(current).parentElement.bgcolor="#ffffff"
End Sub

'------------------------------------------------------------------
' Load the content pane (middle) with the appropriate source file -
'------------------------------------------------------------------
Sub NumClick(myurl)
		
	If sIsIE="4.0>" Then
		Document.All("hotnumber").Item(current).Style.FontWeight="Normal"
		Document.All("hotnumber").Item(current).Style.FontSize="100%"
		Document.All("hotnumber").Item(current).parentElement.bgColor="#ffcc00"

		For q = 0 To Document.All("hotnumber").Length - 1
			If Window.Event.srcElement.Contains(Document.All("hotnumber").Item(q)) Then
				current=q
			End If
		Next

		Document.All("hotnumber").Item(current).Style.fontweight="bold"
		Document.All("hotnumber").Item(current).Style.fontsize="110%"
		Document.All("hotnumber").Item(current).parentElement.bgColor="#ffffff"
	End If
	Top.Frames(1).Location.Href=myurl

End Sub

Sub IE()
	' Empty procedure provided to ensure HREF
	' compatibility between IE 3.02, 4.0x and 5.0
	' It works by George!
End Sub