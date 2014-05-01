%REM
    Agent datastructure
    Created Apr 4, 2014 by Nathan Hilbert/DAI
    Description: Comments for Agent
    Notes: This must be set to All Documents in Database and Agent list selection
%END REM
Option Public
Option Declare
Option Base 1



Dim db As NotesDatabase
Dim doc As NotesDocument
Dim s As NotesSession
Dim tmpName As NotesName
Dim userView As NotesView
Dim userDoc As NotesDocument
Dim grantsView As NotesView
Dim grantsDoc As NotesDocument

'Other variables
Dim i As Integer

Const amp = |&|
Const BR = |<br />|
Const comma = |,|
Const errorStr = |ERROR: |
Const quoteStr = |"|

Const jsonStart = |{|
Const jsonEnd = |}|

Const featureCollectionStart = |{ "type": "FeatureCollection",|
Const featureCollectionEnd = |}|



Const featuresStart = |"features": [|
Const featuresEnd = |]|

Const featureTypeStart = |{ "type": "Feature",|
Const featureTypeEnd = |}|

Const arrayStart = "["
Const arrayEnd = "]"


Const baseURL = |http://localhost:81/SIKANorthTAMIS_Test.nsf/datastructure?|


Dim jsonText As String
Dim newline As String







%REM
    Sub Initialize
    Description: Comments for Sub
%END REM
Sub Initialize()
    'Print out this line to force Domino to not write it"s own 
    'HTML gunk at the beginning of the resulting page
    

    Dim cmdName As String
    Dim queryStr As String
    Dim servicesStr As String  
    Dim tmpStr As String
    Dim tmpInt As Integer
    newline = Chr(10)

    'Initialize our Notes session object
    Set s = New NotesSession
    'Then get a handle to the current database
    Set db = s.CurrentDatabase
    'Get a handle to the agent"s context (header variables and so on)
    Set doc = s.DocumentContext 


    'Parse the command line and call the correct function
    queryStr = doc.Query_String_Decoded(0) & amp
    cmdName = GetCmdLineValue(queryStr, "cmd=", amp)
    'Print |Content-type: application/vnd.ms-excel;|
    'Print |Content-disposition: attachment; filename=data.csv|
    'Print |content-type: text/html;|
    If cmdName = "" Then
        Print "Content-Type:text/html"
        getFormList(servicesStr)
    ElseIf cmdName = "getFields" Then
        Print "Content-Type:text/html"
        getFieldList(queryStr)
    ElseIf cmdName = "showURL" Then
        Print "Content-Type:text/html"
        Print |<h2>Click the following link to download the files</h2>|
        Dim newURL As String
        newURL = Findreplace(queryStr, "showURL", "getValues")
        Print |<a href="| & baseURL & newURL & |">| baseURL & newURL & |</a>|
    ElseIf cmdName = "getValues" Then
        getValueList(queryStr)
    End If


End Sub
%REM
    Sub getFieldList
    Description: Comments for Sub
%END REM
Sub getFieldList(queryString As String)
    Print |<form method="get"><input type="hidden" name="cmd" value="showURL"/> <input type="hidden" name="OpenAgent" value="1"/>|
    Print |<h2>File Type</h2><br>
            <input name="fileformat" type="radio" value="csv" checked="checked">CSV<br>
            <input name="fileformat" type="radio" value="json">JSON<br>
            <input name="fileformat" type="radio" value="geojson">GeoJSON<br>
            <h2>Select Latitude and Longitude (GeoJSON only)<br><br>|
    Dim formselection As String
    formselection = GetCmdLineValue(queryString, "FormOptions=", amp)
    
    If formselection = || Then
        Print |You didn't select anything.  Click back|
    End If
    
    
    
    
    Dim optionsname As String 
    optionsname = |FieldOptions|
    
    Dim session As New NotesSession 

    Dim db As NotesDatabase 

    Dim collection As NotesDocumentCollection

    Dim doc As NotesDocument
    

    Set db = session.CurrentDatabase

    Set collection = db.AllDocuments

    Set doc = collection.GetFirstDocument()
    Dim formTypeList List As Double
    'empList("Maria Jones") = 12345
    'If IsElement(empList(ans$)) = True then
    Dim counter As String
    
    Dim includeThis As Integer
    Dim finalCounter List As Integer

    While Not(doc Is Nothing)
        counter = 0  
        includeThis = 0
        ForAll i In doc.Items 
            'Print i.Name & |<br>|
            If i.Name = |Form| And i.Text = formselection Then
                includeThis = 1
                Exit ForAll
            End If
        End ForAll
        
        If includeThis = 1 Then
            'Print |Form: | & i.Text &|<br>|
            ForAll i In doc.Items
                If IsElement(finalCounter(i.Name)) Then
                    finalCounter(i.Name) = finalCounter(i.Name) + 1
                Else
                    finalCounter(i.Name) = 1
                End If
            End ForAll
        End If

        Set doc = collection.GetNextDocument(doc) 
    Wend
    
    
    
    Dim optionsstring As String
    optionsstring = |<option value=""></option>|
    
    ForAll formType In finalCounter
        optionsstring = optionsstring & |<option value="| & ListTag(formType) & |">| & ListTag(formType) & |</option>|
    End ForAll
    Print |Latitude: | 
    Print |<select name='lat'>|
    Print optionsstring
    Print |</select><br>|
    
    Print |Longitude: |
    Print |<select name='lon'>|
    Print optionsstring
    Print |</select><br><h2>Select Field Values</h2>|
    
    
    ForAll formType In finalCounter
        Print PrintChoice(optionsname, ListTag(formType), ListTag(formType) & | (| & formType & |)|)
    End ForAll
    
    Print |<input type="submit" value="Submit">|
End Sub


Function LSescape(strIn As String) As String
'
' This function performs the equivalent of a JavaScript escape.
' Kenneth H?man, TJ Group AB.
'
Dim strAllowed As String
Dim i As Integer
Dim strChar As String
Dim strReturn As String

'These are the characters that the JavaScript escape-function allows, so we let them pass
'unchanged in this function as well.
strAllowed = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 " & "@/.*-_"
i = 1
strReturn = ""

While Not (i > Len(strIn))
strChar = Mid$(strIn, i, 1)
If InStr(1, strAllowed, strChar) > 0 Then
strReturn = strReturn & strChar
Else
strReturn = strReturn & "%" & Hex$(Asc(strChar))
End If
i = i + 1
Wend

LSescape = strReturn

End Function
%REM
    Function printGeoJSON
    Description: Comments for Function
%END REM
Function printGeoJSON(queryStr As String)
    
    
    Dim optionList List As String
    Dim optionsname As String 
    optionsname = |FieldOptions|
    Call GetCmdLineList(queryStr, optionsname, optionList)
    Dim outputtext As String
    Dim isfirst As Integer
    isfirst =1
    
    ForAll z In optionList
        If isfirst =1 Then
            outputtext = outputtext & z
            isfirst = 0
        Else
            outputtext = outputtext & |,| & z
        End If
    End ForAll
    outputtext = outputtext & Chr(10)
    
    
    Dim session As New NotesSession 

    Dim db As NotesDatabase 

    Dim collection As NotesDocumentCollection

    Dim doc As NotesDocument
    

    Set db = session.CurrentDatabase

    Set collection = db.AllDocuments

    Set doc = collection.GetFirstDocument()
    
    Dim outputLineList List As Variant
    
    Dim entryinlistcheck As Integer
    
    Dim isfirstFeature As Integer
    isfirstFeature = 1
    
    jsonText = featureCollectionStart
    jsonText = jsonText & featuresStart
    
    Dim latField As String
    Dim lonField As String
    latField = GetCmdLineValue(queryStr, "lat=", amp)
    lonField = GetCmdLineValue(queryStr, "lon=", amp)

    While Not(doc Is Nothing)
        
        'print the cooridnates first
        
        
        
        Erase outputLineList
        ForAll i In doc.Items 
            'Print i.Name & |<br>|
            entryinlistcheck = EntryInList(i.Name, optionList)
            If entryinlistcheck > 1 Then
                outputLineList(i.Name) = LSescape(i.Text)
            End If
            If i.Name = latField Then
                outputLineList("latitudeJSON") = Findreplace(i.Text,",",".")
            ElseIf i.Name = lonField Then
                outputLineList("longitudeJSON") = Findreplace(i.Text,",",".")
            End If

        End ForAll
        
        Dim latstring As String
        latstring = outputLineList("latitudeJSON")
        Dim lonstring As String
        lonstring = outputLineList("longitudeJSON")
        
        If CountCharacters(latstring, ".") = 1 And CountCharacters(lonstring, ".") = 1 Then

            If isfirstFeature = 1 Then
                jsonText = jsonText & featureTypeStart
                isfirstFeature  =0
            Else
                jsonText = jsonText & "," & featureTypeStart
            End If
            
            jsonText = jsonText & |"geometry": {"type": "Point", "coordinates": [| & _
            outputLineList("longitudeJSON") & |,| & outputLineList("latitudeJSON") & |]}, "properties":{| 
            
            isfirst = 1
            
    
            ForAll x In optionList
    
                If isfirst =1 Then
                    If IsElement(outputLineList(x)) Then
                        jsonText = jsonText & |"| & x & |":"| & outputLineList(x) & |"|
                    Else
                        jsonText = jsonText & |"| & x & |":""|
                    End If
                    isfirst = 0
                Else
                    If IsElement(outputLineList(x)) Then
                        jsonText = jsonText & |,"| & x & |":"| & outputLineList(x) & |"|
                    Else
                        jsonText = jsonText & |,"| & x & |":""|
                    End If
                    
                End If
            End ForAll
            'end the properties set
            jsonText = jsonText & |}|
            'end the feature
            jsonText = jsonText & featureTypeEnd
        
        End If

        Set doc = collection.GetNextDocument(doc) 
    Wend
    jsonText = jsonText & featuresEnd
    Print jsonText & featureCollectionEnd

End Function




Function PrintRadio(thename As String, thevalue As String, thelabel As String) As String
    PrintRadio = |<input name="| &thename & |" type="radio" value="| & thevalue & |">| & thelabel &|<br>|
End Function




Function PrintChoice(thename As String, thevalue As String, thelabel As String) As String
    PrintChoice = |<input name="| &thename & |" type="checkbox" value="| & thevalue & |">| & thelabel &|<br>|
End Function
%REM
    Function printJSON
    Description: Comments for Function
%END REM
Function printJSON(queryStr As String)
    
    Dim optionList List As String
    Dim optionsname As String 
    optionsname = |FieldOptions|
    Call GetCmdLineList(queryStr, optionsname, optionList)
    Dim outputtext As String
    Dim isfirst As Integer
    isfirst =1
    
    ForAll z In optionList
        If isfirst =1 Then
            outputtext = outputtext & z
            isfirst = 0
        Else
            outputtext = outputtext & |,| & z
        End If
    End ForAll
    outputtext = outputtext & Chr(10)
    
    
    Dim session As New NotesSession 

    Dim db As NotesDatabase 

    Dim collection As NotesDocumentCollection

    Dim doc As NotesDocument
    

    Set db = session.CurrentDatabase

    Set collection = db.AllDocuments

    Set doc = collection.GetFirstDocument()
    
    Dim outputLineList List As Variant
    
    Dim entryinlistcheck As Integer
    
    Dim isfirstFeature As Integer
    isfirstFeature = 1
    
    jsonText = |[|

    While Not(doc Is Nothing)
        Erase outputLineList
        ForAll i In doc.Items 
            'Print i.Name & |<br>|
            entryinlistcheck = EntryInList(i.Name, optionList)
            If entryinlistcheck > 1 Then
                outputLineList(i.Name) = LSescape(i.Text)
            End If

        End ForAll
        
        If isfirstFeature = 1 Then
            jsonText = jsonText & "{"
            isfirstFeature  =0
        Else
            jsonText = jsonText & ",{"
        End If
        
        isfirst = 1
        

        ForAll x In optionList

            If isfirst =1 Then
                If IsElement(outputLineList(x)) Then
                    jsonText = jsonText & |"| & x & |":"| & outputLineList(x) & |"|
                Else
                    jsonText = jsonText & |"| & x & |":""|
                End If
                isfirst = 0
            Else
                If IsElement(outputLineList(x)) Then
                    jsonText = jsonText & |,"| & x & |":"| & outputLineList(x) & |"|
                Else
                    jsonText = jsonText & |,"| & x & |":""|
                End If
                
            End If
        End ForAll
        jsonText = jsonText & |}|

        Set doc = collection.GetNextDocument(doc) 
    Wend
    Print jsonText & |]|    
        
End Function





Function EntryInList (Value As Variant, ValueList As Variant) As Integer
    ' This will return a 1 based value if the position in the list
    EntryInList = 0
    Dim zi As Integer
    zi = 1
    ForAll Entries In ValueList
        If Entries = Value Then
            EntryInList = zi
            Exit Function
        End If
        zi = zi + 1
    End ForAll
End Function



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
        tmpInt = InStr( tempstring, optionitem)
        If (tmpInt > 0) Then
            tempsplitarray = Split(tempstring, |=|)
            optionList(tempsplitarray(1)) = tempsplitarray(1)
        End If
        
        counter = counter + 1
    Loop

End Function




Sub getFormList(servicesStr As String)
    Print |<form method="get"><input type="hidden" name="cmd" value="getFields"/> <input type="hidden" name="OpenAgent" value="1"/>|
    Dim optionsname As String 
    optionsname = |FormOptions|
    
    Dim session As New NotesSession 

    Dim db As NotesDatabase 

    Dim collection As NotesDocumentCollection

    Dim doc As NotesDocument
    

    Set db = session.CurrentDatabase

    Set collection = db.AllDocuments

    Set doc = collection.GetFirstDocument()
    Dim formTypeList List As Double
    'empList("Maria Jones") = 12345
    'If IsElement(empList(ans$)) = True then
    Dim counter As String

    While Not(doc Is Nothing)
        counter = 0  
        
        ForAll i In doc.Items 
            'Print i.Name & |<br>|
            If i.Name = "Form" Then
                'Print |Form: | & i.Text &|<br>|
                If IsElement(formTypeList(i.Text)) Then
                    formTypeList(i.Text) = formTypeList(i.Text) + 1
                Else
                    formTypeList(i.Text) = 1
                End If
            End If
            If Not i.Text = ""  Then
                counter = counter + 1
            End If
            'strFields = strFields & Chr(10) & i.Name & " = " & i.Text 
        End ForAll

        Set doc = collection.GetNextDocument(doc) 
    Wend
    
    ForAll formType In formTypeList
        
        Print PrintRadio(optionsname, ListTag(formType), ListTag(formType) & | (| & formType & |)|)
    End ForAll
    
    Print |<input type="submit" value="Submit">|

End Sub







Function CountCharacters(searchstring As String, searchfor As String) As Integer
     Dim count As Integer
     count = 0
     Do While InStr(searchstring, searchfor) > 0
          count = count + 1
          searchstring = StrRight(searchstring, searchfor)
     Loop
     CountCharacters = count
End Function
%REM
    Function printCSV
    Description: Comments for Function
%END REM
Function printCSV(queryStr As String)
    Dim optionList List As String
    Dim optionsname As String 
    optionsname = |FieldOptions|
    Call GetCmdLineList(queryStr, optionsname, optionList)
    Dim outputtext As String
    Dim isfirst As Integer
    isfirst =1
    ForAll z In optionList
        If isfirst =1 Then
            outputtext = outputtext & z
            isfirst = 0
        Else
            outputtext = outputtext & |,| & z
        End If
    End ForAll
    outputtext = outputtext & Chr(10)
    
    
    Dim session As New NotesSession 

    Dim db As NotesDatabase 

    Dim collection As NotesDocumentCollection

    Dim doc As NotesDocument
    

    Set db = session.CurrentDatabase

    Set collection = db.AllDocuments

    Set doc = collection.GetFirstDocument()
    
    Dim outputLineList List As Variant
    
    Dim entryinlistcheck As Integer
    
    

    While Not(doc Is Nothing)
        Erase outputLineList
        ForAll i In doc.Items 
            'Print i.Name & |<br>|
            entryinlistcheck = EntryInList(i.Name, optionList)
            If entryinlistcheck > 1 Then
                outputLineList(i.Name) = LSescape(i.Text)
            End If

        End ForAll
        isfirst = 1

        ForAll x In optionList

            If isfirst =1 Then
                If IsElement(outputLineList(x)) Then
                    outputtext = outputtext & outputLineList(x)
                End If
                isfirst = 0
            Else
                If IsElement(outputLineList(x)) Then
                    outputtext = outputtext & |,| & outputLineList(x)
                Else
                    outputtext = outputtext & |,|
                End If
                
            End If
        End ForAll
        outputtext = outputtext & Chr(10)

        Set doc = collection.GetNextDocument(doc) 
    Wend
    Print outputtext
    
    
End Function



Function GetCmdLineValue( textStr As String, delim1 As String, delim2 As String) As String

  Dim startPos As Integer  
  Dim tmpInt As Integer
  Dim valLen As Integer

  'find the first ocurrance of the delimeter
  tmpInt = InStr( textStr, delim1)
  'Only continue if we"ve found something
  If (tmpInt > 0) Then
    'Figure out where the value starts 
    startPos = tmpInt + Len(delim1)
    'Then look past there for the second delimeter
    valLen = InStr(startPos, textStr, delim2) - startPos
    'The value we"re looking for is between the two delimeters
    GetCmdLineValue = Mid( textStr, startPos, valLen)
  Else
    GetCmdLineValue = ||
  End If   
End Function

Sub getValueList(queryStr As String)
    Dim fileFormat As String
    fileFormat = GetCmdLineValue(queryStr, |fileformat=|, |&|)
    If fileFormat = "json" Then
        Print |Content-Type: application/json|
        printJSON(queryStr)
    ElseIf fileFormat = "geojson" Then
        Dim lat As String
        Dim lon As String
        lat = GetCmdLineValue(queryStr, "lat=", amp)
        lon = GetCmdLineValue(queryStr, "lon=", amp)
        If lat = "" Or lon = "" Then
            Print |Content-Type: text/html|
            Print |Please select a valid lat and lon for the GeoJSON format.  Click back.|
        Else
            Print |Content-Type: application/json|
            printGeoJSON(queryStr)
        End If
        
    Else
        Print |Content-type: application/vnd.ms-excel;|
        Print |Content-disposition: attachment; filename=data.csv|
        'Print |Content-type: text/html|
        printCSV(queryStr)
    End If


End Sub






Function Findreplace(ByVal wholestring As Variant, find As String, ireplace As String) As String
    Dim checkstring As String
    Dim saveleft As String
    Dim n As Integer
    Dim leftstring As String
    Dim rightstring As String
    checkstring=wholestring
    saveleft=""
    While InStr(1, checkstring, find)<>0 
        n=InStr(1, checkstring, find)
        leftstring = Left(checkstring, n-1)
        rightstring=Right(checkstring, Len(checkstring)-n-Len(find)+1)
        saveleft=saveleft+leftstring+ireplace
        checkstring=rightstring
    Wend
    FindReplace= saveleft+checkstring
End Function

