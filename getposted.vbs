%REM
    Agent getposted
    Created Apr 4, 2014 by Nathan Hilbert/DAI
    Description: Comments for Agent
%END REM
Option Public
Option Declare


Sub Initialize
    Dim session As New NotesSession
    Dim doc As NotesDocument
    Set doc=session.DocumentContext
    Dim req As String
    req=doc.Request_Content(0)
    Print Len(req)
    '   For x=Len(req) To 1 Step -1
    '       tmp=tmp+Mid(req,x,1)
    '   Next
    Print req
    
    Dim optionList List As String
    
    Call GetCmdLineList(req, |mychoices|, optionList)
    
    ForAll theoption In optionList
        Print |theoption| & theoption & |<br>|
    End ForAll
    
    
    
    Print |<form method="post">
  <input name="mychoices" type="checkbox" value="volvo">Volvo
  <input name="mychoices" type="checkbox" value="saab">Saab
  <input name="mychoices" type="checkbox" value="opel">Opel
  <input name="mychoices" type="checkbox" value="audi">Audi
    <input type="text" name="something"><input type="submit" value="Submit"></form>|
End Sub




Function GetCmdLineList( textStr As String, optionitem As String, optionList List As String)
    Dim splitarray As Variant
    Dim tempsplitarray As Variant
    splitarray = Split(textStr, "&")
    Dim counter As Integer
    counter = 0
    Dim upperbound As Integer
    Dim tmpInt As Integer
    
    upperbound = UBound(splitarray)

    Do While counter <= upperbound
        Dim tempstring As String
        tempstring = splitarray(counter)
        Print tempstring
        tmpInt = InStr( tempstring, optionitem)
        Print tmpInt
        If (tmpInt > 0) Then
            tempsplitarray = Split(tempstring, |=|)
            optionList(tempsplitarray(1)) = tempsplitarray(1)
        End If
        
        counter = counter + 1
    Loop

End Function
%REM
    Function PrintChoice
    Description: Comments for Function
%END REM
Function PrintChoice(thename As String, thevalue As String, thelabel As String) As String
    PrintChoice = |<input name="mychoices" type="checkbox" value="audi">Audi|
End Function