VERSION 5.00
Object = "{2668C1EA-1D34-42E2-B89F-6B92F3FF627B}#5.0#0"; "scivb2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   13470
   StartUpPosition =   2  'CenterScreen
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'todo: if only one table is selected from do not require table name for field auto list

Dim cn As New Connection

Private intellisense As New Collection
Private tableList As String

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
    
    cn.ConnectionString = "Provider=MSDASQL;Driver={Microsoft Access Driver (*.mdb)};DBQ=" & App.Path & "\db1.mdb;"
    cn.Open
 
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
    Dim prev As String, absPos As Long, tmp As String
    Dim it As CIntellisenseItem
    
    tmp = getTextToCurPos(absPos)
    If isInQuotedString(tmp, absPos) Then Exit Sub

    prev = sci.PreviousWord
    
    For Each it In intellisense
        If LCase(className) = LCase(it.objName) Or LCase(prev) = LCase(it.objName) Then
            sci.ShowAutoComplete it.methods
            Exit Sub
        End If
    Next
    
    For Each it In intellisense
        If LCase(className) = LCase(Left(it.objName, Len(className))) Or LCase(prev) = LCase(Left(it.objName, Len(prev))) Then
            sci.ShowAutoComplete Trim(tableList)
            Exit Sub
        End If
    Next
            
End Sub

Private Sub sci_KeyUp(KeyCode As Long, Shift As Long)
    Dim prev As String, tables As String
    Dim it As CIntellisenseItem
    Dim showTableList As Boolean, fragment As String, absPos As Long, tmp As String
    
    If KeyCode < 32 Or KeyCode > 127 Then Exit Sub 'abort if not printable
    If sci.DirectSCI.AutoCActive Then Exit Sub 'if the auto complete is active maybe arrow up/down
   
    'really hate doing this on every keypress :(
    tmp = getTextToCurPos(absPos)
    If isInQuotedString(tmp, absPos) Then Exit Sub
     
    prev = LCase(prevKeyWord(fragment))
    
    If prev = "from" Or prev = "join" Then
        If Len(fragment) = 0 Or Right(Trim(fragment), 1) = "," Then showTableList = True
    End If
    
    If showTableList Then
        For Each it In intellisense
            'we do not want to autocomplete a table name fragement here on keyup to invasive
            If Len(fragment) > 0 Then
                If LCase(VBA.Left(it.objName, Len(fragment))) = LCase(fragment) Then Exit Sub
            End If
            tables = tables & it.objName & " "
        Next
        sci.ShowAutoComplete Trim(tables)
    End If
    
End Sub

Function getTextToCurPos(Optional ByRef absPos As Long) As String
    Dim tmp As String, linePos As Long, lineStartPos As Long
    On Error GoTo hell
    tmp = sci.Text
    linePos = sci.GetCaretInLine()
    lineStartPos = sci.PositionFromLine(sci.CurrentLine)
    absPos = lineStartPos + linePos
    getTextToCurPos = Mid(tmp, 1, absPos)
hell:
End Function


Function prevKeyWord(ByRef lastFragment As String) As String
    Dim tmp As String, linePos As Long, lineStartPos As Long, absPos As Long
    Dim ary() As String, keys() As String, i As Long, j As Long, lastwordStartPos As Long
    
    On Error GoTo hell

    tmp = getTextToCurPos()
    tmp = Replace(Replace(tmp, vbCrLf, Empty), vbTab, Empty) 'remove tabs, new lines and spaces
    tmp = removeQuotedStrings(tmp)
    
    ary() = Split(tmp, " ")
    keys = Split(keywords, " ")
    lastwordStartPos = Len(tmp)
    
    For i = UBound(ary) To 0 Step -1
        For j = 0 To UBound(keys)
            If ary(i) = keys(j) Then
                prevKeyWord = keys(j)
                lastFragment = getLastFragment(tmp, lastwordStartPos)
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
    tableList = tableList & className & " "
    it.objName = className
    it.methods = spaceSeperatedMethodList
    intellisense.Add it
    AddIntellisense = True
    
End Function


