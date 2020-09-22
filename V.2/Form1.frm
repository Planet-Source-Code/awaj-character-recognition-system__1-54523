VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Character Recognizer"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdLoadFont 
      Caption         =   "Load Font"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2640
      Top             =   600
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   0
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "FontInfo"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdIdentifyChar 
      Caption         =   "Identify Character"
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdLoadData 
      Caption         =   "Load Image Data"
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' first byte is blue
' 2nd byte is green
' 3rd byte is red
Private Type Header
    'BMP Header
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
    'file Info
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type fntSpec
    fontChar As String
    charDim As String
    charInfo As String
End Type

Private BMPHeader As Header
Private bAlignFac As Integer ' each row must be multiple of 4
Private color(1 To 3) As Byte
Private dataGrid() As Byte
Private dataGridTemp() As Byte
Private fontSpec() As fntSpec
Private diceGrid As String
Private outPut As String

Private Sub cmdIdentifyChar_Click()
    Dim i As Long, j As Long
    Dim start As Variant, finish As Variant, totalTime As Variant
    cmdIdentifyChar.Enabled = False
    start = Timer
    outPut = ""
    ProgressBar2.Min = 0
    ProgressBar2.Max = UBound(dataGrid, 1) + 1
    ProgressBar2.Value = 0
    For i = UBound(dataGrid, 1) To 0 Step -1
        For j = 0 To UBound(dataGrid, 2)
            If dataGrid(i, j) <> 255 Then
                j = j + findChar(i, j)
                DoEvents
            End If
        Next j
        ProgressBar2.Value = ProgressBar2.Value + 1
    Next i
    finish = Timer
    totalTime = finish - start
    MsgBox "Total processing Time: " & totalTime & " (s)"
    MsgBox outPut
    cmdLoadData.Enabled = True
End Sub

Private Function findChar(ByVal indi As Long, ByVal indj As Long) As Long
    Dim i As Long, j As Long, k As Long
    Dim temp As Variant
    Dim temp1 As Long, temp2 As Long
    
    For k = 0 To UBound(fontSpec)
        temp = Split(fontSpec(k).charDim, "*")
        diceGrid = ""
        temp1 = indi - (temp(0) - 1)
        If indi - (temp(0) - 1) >= 0 And (UBound(dataGrid, 2) - indj) >= (temp(1) - 1) Then
            For i = indi To indi - (temp(0) - 1) Step -1
                For j = indj To indj + (temp(1) - 1)
                    diceGrid = diceGrid & dataGrid(i, j) & ","
                Next j
            Next i
            If diceGrid = fontSpec(k).charInfo Then
                outPut = outPut & fontSpec(k).fontChar & ","
                For i = indi To indi - (temp(0) - 1) Step -1
                    For j = indj To indj + (temp(1) - 1)
                        dataGrid(i, j) = 255
                    Next j
                Next i
                findChar = temp(1) - 1
                Exit Function
            End If
        End If
    Next k
    findChar = 0
End Function

Private Sub cmdLoadData_Click()
    Dim i As Long, rowEnd As Long, a As Long, b As Long
    Dim colorCounter As Integer
    Dim dataI As Long, dataJ As Long
    Dim hasRow As Boolean
    
    cmdLoadData.Enabled = False
    
    'Read File Header
    Open App.Path & "\untitled.bmp" For Random As #1 Len = Len(BMPHeader)
    Get #1, 1, BMPHeader
    If BMPHeader.bfType = 19778 Then
        'MsgBox BMPHeader.biSizeImage
        'ReDim dataGrid(BMPHeader.biHeight - 1, BMPHeader.biWidth - 1)
    Else
        MsgBox "File format is not '*.bmp'"
    End If
    Close #1
    '************************************************'
    ProgressBar1.Min = 0
    ProgressBar1.Max = BMPHeader.biHeight
    ProgressBar1.Value = 0
    Open App.Path & "\untitled.bmp" For Random Access Read Write As #1 Len = 1
    Seek 1, 55
    
    i = 1
    rowEnd = BMPHeader.biWidth * 3
    colorCounter = 1
    hasRow = False
    dataI = -1
    dataJ = 0
    While i <= BMPHeader.biSizeImage
        'read data
        Get #1, , color(colorCounter)
        
        If colorCounter = 3 Then
            'do processing
            If color(1) = 0 And color(2) = 0 And color(3) = 0 Then
                If hasRow = False Then
                    dataI = dataI + 1
                    ReDim Preserve dataGridTemp(BMPHeader.biWidth - 1, dataI)
                    dataGridTemp(dataJ, dataI) = 1
                    hasRow = True
                Else
                    dataGridTemp(dataJ, dataI) = 1
                End If
            End If
            dataJ = dataJ + 1
            colorCounter = 0
        End If
        'check end of row
        If i = rowEnd Then
            ProgressBar1.Value = ProgressBar1.Value + 1
            dataJ = 0
            hasRow = False
            DoEvents
            Select Case i Mod 4
            Case 1
                bAlignFac = 3
            Case 2
                bAlignFac = 2
            Case 3
                bAlignFac = 1
            Case 0
                bAlignFac = 0
            End Select
            If bAlignFac > 0 Then
                Seek 1, (Loc(1) + bAlignFac + 1)
                i = i + bAlignFac
            End If
            rowEnd = rowEnd + bAlignFac + BMPHeader.biWidth * 3
        End If
        colorCounter = colorCounter + 1
        i = i + 1
    Wend
    Close #1
    
    dataI = -1
    For i = 0 To UBound(dataGridTemp, 1)
        For dataJ = UBound(dataGridTemp, 2) To 0 Step -1
            If dataGridTemp(i, dataJ) = 1 Then
                dataI = dataI + 1
                dataJ = UBound(dataGridTemp, 2)
                ReDim Preserve dataGrid(UBound(dataGridTemp, 2), dataI)
                For a = 0 To UBound(dataGridTemp, 2)
                    dataGrid(a, dataI) = dataGridTemp(i, dataJ)
                    dataJ = dataJ - 1
                Next a
            End If
        Next dataJ
    Next i
    Erase dataGridTemp
    DoEvents
    MsgBox "Done"
    cmdIdentifyChar.Enabled = True
End Sub

Private Sub cmdLoadFont_Click()
    Dim i As Long
    cmdLoadFont.Enabled = False
    cmdLoadData.Enabled = True
    ReDim fontSpec(Adodc1.Recordset.RecordCount - 1)
    i = 0
    Adodc1.Recordset.MoveFirst
    While Not Adodc1.Recordset.EOF
        fontSpec(i).fontChar = Adodc1.Recordset.Fields("FontChar")
        fontSpec(i).charDim = Adodc1.Recordset.Fields("CharDimention")
        fontSpec(i).charInfo = Adodc1.Recordset.Fields("CharInfo")
        Adodc1.Recordset.MoveNext
        i = i + 1
    Wend
    MsgBox "done"
End Sub

Private Sub Form_Load()
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\FontInfo.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
    Adodc1.CommandType = adCmdTable
    Adodc1.RecordSource = "FontInfo"
    Adodc1.Refresh
End Sub
