VERSION 5.00
Object = "{2668C1EA-1D34-42E2-B89F-6B92F3FF627B}#5.0#0"; "scivb2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   13470
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboTest 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   810
      TabIndex        =   3
      Top             =   3735
      Width           =   12615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   45
      Width           =   3165
   End
   Begin sci2.SciSimple sci 
      Height          =   3615
      Left            =   3375
      TabIndex        =   0
      Top             =   0
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   6376
   End
   Begin VB.Label Label1 
      Caption         =   "Tests"
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   3735
      Width           =   510
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'todo: if only one table is selected from do not require table name for field auto list

'http://proton-ce.sourceforge.net/rc/scintilla/doku/ScintillaDoc.html

Dim cn As New Connection

Private intellisense As New Collection 'currently loaded with all tables, all fields once at startup
Private tableList As String
Private lastKeycode As Long
Private lastShift As Long
Private lastSize As Long

Private Sub cboTest_Click()
    sci.Text = cboTest.Text 'triggers sci_OnModified
End Sub

Private Sub sci_OnModified(Position As Long, modificationType As Long)
    If Abs(lastSize - sci.DirectSCI.GetLength) > 8 Then
         cleanExtractAlias 'they pasted something or its a new document
    End If
    lastSize = sci.DirectSCI.GetLength
End Sub

'this happens on a full reparse, note: test mode can not discern from a table alias and a field alias
Function cleanExtractAlias(Optional test As String = Empty)
    On Error Resume Next
    
    Dim it As CIntellisenseItem, base As String, tmp() As String, prev As String, nxt As String, i As Long
    Dim testMode As Boolean 'do not require database driven intellisense to parse
    
    If Len(test) = 0 Then
        test = sci.Text
        If intellisense.Count = 0 Then Exit Function
        For Each it In intellisense
            it.alias = Empty
        Next
    Else
        Debug.Print "testing cleanExtractAlias: " & test
        testMode = True
    End If
    
    Debug.Print "alias reparse " & Now
    base = Replace(Replace(test, vbCrLf, " "), vbTab, " ")
    base = LCase(removeQuotedStrings(base))
    
    While InStr(base, "  ") > 0
        base = Replace(base, "  ", " ")
    Wend
    
    tmp = Split(base, " ")
    For i = 0 To UBound(tmp)
        If tmp(i) = "as" Then
             prev = tmp(i - 1)
             nxt = Replace(tmp(i + 1), ",", Empty)
             If testMode Then
                Debug.Print "   alias: " & prev & " -> " & nxt
             Else
                If findIT(prev, it) Then
                   it.alias = nxt
                   Debug.Print "alias: " & it.objName & " -> " & it.alias
                End If
            End If
        End If
    Next
    
End Function

 Function findIT(ByVal name As String, ByRef outIT As CIntellisenseItem) As Boolean
    Dim it As CIntellisenseItem
    name = LCase(name)
    Set outIT = Nothing
    For Each it In intellisense
        If it.objName = name Then
            Set outIT = it
            findIT = True
            Exit Function
        End If
    Next
 End Function

Private Sub Form_Load()
    On Error Resume Next
    
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim table As String, cols As String
    
    sci.LoadHighlighter App.Path & "\sql.hilighter", True
      
    'AddIntellisense "test", "t_field1 t_field2 t_field3"
    'AddIntellisense "fart", "f_field1 f_field2 f_field3"
    'AddIntellisense "knocker", "k_field1 k_field2 k_field3"
    
    sci.Text = "select * from"
    cboTest.AddItem "SELECT C.ID, C.Title, C.Salary, O.age FROM tblPositions AS C, tblUsers AS O WHERE C.ID = O.ID;"
    cboTest.AddItem "SELECT tblPositions.*, tblUsers.* FROM tblPositions, tblUsers;"
    cboTest.AddItem "SELECT tblPositions.*, tblUsers.* FROM tblPositions LEFT JOIN tblUsers ON tblPositions.id = tblUsers.id"
    cboTest.AddItem "SELECT * FROM tblPositions WHERE id IN (SELECT id FROM tblUsers);"


    cn.ConnectionString = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\db1.mdb;"
    cn.Open
 
    'should we load tables on demand and cache them? need to bench mark once out of test...
    Set rs = cn.OpenSchema(adSchemaTables)

    Do Until rs.EOF
        If rs("TABLE_TYPE") = "TABLE" Then
          table = rs("TABLE_NAME").Value
          cols = getColumns(table)
          AddIntellisense table, cols
        End If
        rs.MoveNext
    Loop
    
    cn.Close
    
    'cleanExtractAlias "SELECT C.ID, C.NAME, C.AGE, O.AMOUNT FROM CUSTOMERS AS C, ORDERS AS O WHERE  C.ID = O.CUSTOMER_ID;"
    'cleanExtractAlias "SELECT t1.*, t2.* FROM t1, t2;"
    'cleanExtractAlias "SELECT t1.*, t2.* FROM t1 LEFT JOIN t2 ON t1.i1 = t2.i2"
    'cleanExtractAlias "SELECT * FROM score WHERE event_id IN (SELECT event_id FROM event WHERE type = 'T');"
    
End Sub

'https://www.vbsedit.com/scripts/data/access/list_tables_columns.asp
Function getColumns(strTableName) As String

  Dim rs2 As Recordset, tmp As String
  
  Set rs2 = cn.OpenSchema(adSchemaColumns, Array(Null, Null, strTableName))

  Do Until rs2.EOF
    tmp = tmp & rs2("COLUMN_NAME") & " "
    rs2.MoveNext
  Loop
  
  getColumns = Trim(tmp)
  rs2.Close
  
End Function

'triggered on ctrl space or obj.
Private Sub sci_AutoCompleteEvent(className As String)
    Dim prev As String, absPos As Long, tmp As String, autoCompleteTableName As Boolean
    Dim it As CIntellisenseItem
    
    tmp = getTextToCurPos(absPos)
    If isInQuotedString(tmp, absPos) Then Exit Sub

    className = LCase(className)
    prev = LCase(sci.PreviousWord)
    
    'ctrl-space
    If lastKeycode = 32 And lastShift = 4 Then
        If Len(className) = 0 And Len(prev) = 0 Then
            sci.ShowAutoComplete Trim(tableList)
            Exit Sub
        End If
    End If
    
    For Each it In intellisense
        If className = it.objName Or prev = it.objName Then
            sci.ShowAutoComplete it.methods
            Exit Sub
        End If
        If Len(it.alias) > 0 Then
            If className = it.alias Or prev = it.alias Then
                sci.ShowAutoComplete it.methods
                Exit Sub
            End If
        End If
    Next
    
    For Each it In intellisense
        If Len(className) > 0 Then If className = Left(it.objName, Len(className)) Then autoCompleteTableName = True
        If Len(prev) > 0 Then If prev = Left(it.objName, Len(prev)) Then autoCompleteTableName = True
        If autoCompleteTableName Then
            sci.ShowAutoComplete Trim(tableList)
            Exit Sub
        End If
    Next
            
End Sub

Private Sub sci_KeyDown(KeyCode As Long, Shift As Long)
    'Debug.Print KeyCode & " " & Shift
    lastKeycode = KeyCode
    lastShift = Shift
End Sub

Private Sub sci_KeyUp(KeyCode As Long, Shift As Long)
    Dim prev As String, tables As String
    Dim it As CIntellisenseItem
    Dim showTableList As Boolean, fragment As String, absPos As Long, tmp As String
    Dim flattened As String, nextBack As String, keywordStartPos As Long
    
    If KeyCode <> 188 Then 'comma
        If KeyCode < 32 Or KeyCode > 127 Then Exit Sub 'abort if not printable
        If sci.DirectSCI.AutoCActive Then Exit Sub 'if the auto complete is active maybe arrow up/down
    End If
    
    'really hate doing this on every keypress :(
    tmp = getTextToCurPos(absPos)
    If isInQuotedString(tmp, absPos) Then Exit Sub
     
    prev = LCase(prevKeyWord(fragment, flattened, nextBack, keywordStartPos))
    
    If prev = "from" Or prev = "join" Then
        If Len(fragment) = 0 Or Right(Trim(fragment), 1) = "," Then showTableList = True
    End If
    
    If prev = "as" Then
        If Len(fragment) = 0 Then Exit Sub                    'they havent entered alias name yet
        If KeyCode <> 32 And KeyCode <> 188 Then Exit Sub     'wait for completion on " " or comma
        If KeyCode = 188 And Right(fragment, 1) = "," Then
            fragment = Mid(fragment, 1, Len(fragment) - 1)
        End If
        If InStr(fragment, ",") > 0 Then Exit Sub 'nope
        For Each it In intellisense
            If it.objName = LCase(nextBack) Then
                it.alias = LCase(Trim(fragment))
                Exit Sub
            End If
        Next
    End If
    
    If showTableList Then
        For Each it In intellisense
            'we do not want to autocomplete a table name fragement here on keyup to invasive
            If Len(fragment) > 0 Then
                If LCase(VBA.Left(it.objName, Len(fragment))) = LCase(fragment) Then Exit Sub
            End If
        Next
        sci.ShowAutoComplete tableList
    End If
    
End Sub

Function getTextToCurPos(Optional ByRef absPos As Long) As String
    Dim tmp As String
    On Error GoTo hell
    tmp = sci.Text
    absPos = sci.DirectSCI.GetCurPos
    getTextToCurPos = Mid(tmp, 1, absPos)
hell:
End Function

'Function getTextFromPos(ByVal absPos As Long) As String
'    On Error GoTo hell
'    getTextFromPos = Mid(sci.Text, absPos)
'hell:
'End Function

''this is a different api design than prev word used for incremental parsing..
'Function nextKeyWord(flattened As String, pos As Long, ByRef endPos As Long) As String
'    Dim tmp As String, linePos As Long, lineStartPos As Long, absPos As Long
'    Dim ary() As String, keys() As String, i As Long, j As Long
'
'    On Error GoTo hell
'    endPos = 0
'
''    nextBack = Empty
''    flattenedStr = Empty
''    lastwordStartPos = 0
'
'    'flattened is assumed to have quoted strings, tabs, new lines removed already
'    ary() = Split(tmp, " ")
'    keys = Split(keywords, " ")
'    absPos = pos
'
'    For i = 0 To UBound(ary)
'        For j = 0 To UBound(keys)
'            If ary(i) = keys(j) Then
'                nextKeyWord = keys(j)
'                endPos = absPos + Len(ary(i))
'                'lastFragment = getLastFragment(tmp, lastwordStartPos)
'                'If i > 0 Then nextBack = ary(i - 1)
'                'Debug.Print """" & lastFragment & """"
'                Exit Function
'            End If
'        Next
'        absPos = absPos + Len(ary(i)) + 1
'    Next
'
'hell:
'End Function

'if are going to do intensive work..lets be able to cache it..
Function prevKeyWord(ByRef lastFragment As String, _
                    Optional ByRef flattenedStr As String, _
                    Optional ByRef nextBack As String, _
                    Optional ByRef lastwordStartPos As Long _
) As String
    Dim tmp As String, linePos As Long, lineStartPos As Long, absPos As Long
    Dim ary() As String, keys() As String, i As Long, j As Long
    
    On Error GoTo hell
    
    nextBack = Empty
    flattenedStr = Empty
    lastwordStartPos = 0
    
    tmp = getTextToCurPos()
    tmp = Replace(Replace(tmp, vbCrLf, Empty), vbTab, Empty) 'remove tabs, new lines and spaces
    tmp = LCase(removeQuotedStrings(tmp))
    flattenedStr = tmp
    
    ary() = Split(tmp, " ")
    keys = Split(keywords, " ")
    lastwordStartPos = Len(tmp)
    
    For i = UBound(ary) To 0 Step -1
        For j = 0 To UBound(keys)
            If ary(i) = keys(j) Then
                prevKeyWord = keys(j)
                lastFragment = getLastFragment(tmp, lastwordStartPos)
                If i > 0 Then nextBack = ary(i - 1)
                'Debug.Print """" & lastFragment & """"
                Exit Function
            End If
        Next
        lastwordStartPos = lastwordStartPos - Len(ary(i)) - 1
    Next
 
hell:
End Function

Function getLastFragment(base As String, endPos As Long) As String
    On Error Resume Next
    getLastFragment = Trim(Mid(base, endPos + 1))
End Function

'tests:
    'Debug.Print removeQuotedStrings("this is 'my string' and i know")
    'Debug.Print removeQuotedStrings("this is '""my string""' and i know")
    'Debug.Print removeQuotedStrings("this is 'my \'string")
    'Debug.Print removeQuotedStrings("this is 'my ""string")
Function removeQuotedStrings(ByVal tmp As String) As String

    Dim t2 As String, inSingleQuote As Boolean, inDoubleQuote As Boolean, c As Byte, prev As Byte
    Dim singleQuoteStart As Long, doubleQuoteStart As Long, i As Long

    If InStr(tmp, "'") > 0 Or InStr(tmp, """") > 0 Then
rescan:
        For i = 1 To Len(tmp)
            c = Asc(Mid(tmp, i, 1))
            
            If c = Asc("'") Then
                If inSingleQuote Then
                    If prev <> Asc("\") Then
                        inSingleQuote = False
                        tmp = Mid(tmp, 1, singleQuoteStart - 1) & " " & Mid(tmp, i + 1)
                        GoTo rescan
                    End If
                Else
                    If Not inDoubleQuote Then
                        inSingleQuote = True
                        singleQuoteStart = i
                    End If
                End If
            End If
            
            If c = Asc("""") Then
                If inDoubleQuote Then
                    If prev <> Asc("\") Then
                        inDoubleQuote = False
                        tmp = Mid(tmp, 1, doubleQuoteStart - 1) & " " & Mid(tmp, i + 1)
                        GoTo rescan
                    End If
                Else
                    If Not inSingleQuote Then
                        inDoubleQuote = True
                        doubleQuoteStart = i
                    End If
                End If
            End If
        
            prev = c
            
        Next
        
        If inDoubleQuote Then 'it wasnt done yet, truncate
            tmp = Mid(tmp, 1, doubleQuoteStart - 1)
        End If
        
        If inSingleQuote Then 'it wasnt done yet, truncate
            tmp = Mid(tmp, 1, singleQuoteStart - 1)
        End If
        
        
    End If
    
    While InStr(tmp, "  ") > 0
        tmp = Replace(tmp, "  ", " ")
    Wend
    
    removeQuotedStrings = tmp
    
End Function

''tests:
'  Debug.Print isInQuotedString("this is 'my string' and i know", Len("this is ")) 'false
'  Debug.Print isInQuotedString("this is 'my string' and i know", Len("this is 'my")) 'true
'  Debug.Print isInQuotedString("this is 'my string' and i know", Len("this is 'my string'")) 'false
'  Debug.Print isInQuotedString("this is 'my string' and i know", Len("this is 'my string' ")) 'false
Function isInQuotedString(tmp As String, pos As Long) As Boolean

    Dim t2 As String, inSingleQuote As Boolean, inDoubleQuote As Boolean, c As Byte, prev As Byte
    Dim singleQuoteStart As Long, doubleQuoteStart As Long, i As Long

    If InStr(tmp, "'") < 1 And InStr(tmp, """") < 1 Then
        isInQuotedString = False
        Exit Function
    End If

    For i = 1 To Len(tmp)
        c = Asc(Mid(tmp, i, 1))
        
        If c = Asc("'") Then
            If inSingleQuote Then
                If prev <> Asc("\") Then
                    inSingleQuote = False
                End If
            Else
                If Not inDoubleQuote Then
                    inSingleQuote = True
                    singleQuoteStart = i
                End If
            End If
        End If
        
        If c = Asc("""") Then
            If inDoubleQuote Then
                If prev <> Asc("\") Then
                    inDoubleQuote = False
                End If
            Else
                If Not inSingleQuote Then
                    inDoubleQuote = True
                    doubleQuoteStart = i
                End If
            End If
        End If
    
        prev = c
        If pos = i Then Exit For
             
    Next
    
    If inDoubleQuote Or inSingleQuote Then
        isInQuotedString = True
    End If
    
End Function

Function AddIntellisense(className As String, ByVal spaceSeperatedMethodList As String) As Boolean
    
    If Len(className) = 0 Or InStr(className, " ") > 1 Then Exit Function
    If Len(spaceSeperatedMethodList) = 0 Then Exit Function
    
    If InStr(spaceSeperatedMethodList, ",") > 0 Then
        spaceSeperatedMethodList = Join(Split(spaceSeperatedMethodList, ","), " ")
    End If
    
    Dim it As CIntellisenseItem
    
    For Each it In intellisense
        If it.objName = className Then Exit Function
    Next
    
    Set it = New CIntellisenseItem
    tableList = tableList & className & " "  'preserve case here..
    it.raw_objName = className
    it.methods = spaceSeperatedMethodList
    intellisense.Add it
    AddIntellisense = True
    
End Function

Function IsIde() As Boolean
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function
