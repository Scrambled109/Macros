VERSION 5.00
Begin VB.UserForm ModifiedPartsListMapper
   Caption         =   "Map Modified Parts List Columns"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7440
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Caption         =   "Cancel"
      Height          =   390
      Left            =   5520
      TabIndex        =   9
      Top             =   3300
      Width           =   1455
   End
   Begin VB.CommandButton cmdRun
      Caption         =   "Apply Properties"
      Default         =   -1  'True
      Height          =   390
      Left            =   3720
      TabIndex        =   8
      Top             =   3300
      Width           =   1695
   End
   Begin VB.ComboBox cboRawMaterial
      Height          =   315
      Left            =   2520
      Style           =   2  'fmStyleDropDownList
      TabIndex        =   7
      Top             =   2460
      Width           =   4455
   End
   Begin VB.ComboBox cboDescription
      Height          =   315
      Left            =   2520
      Style           =   2  'fmStyleDropDownList
      TabIndex        =   5
      Top             =   1860
      Width           =   4455
   End
   Begin VB.ComboBox cboPartNumber
      Height          =   315
      Left            =   2520
      Style           =   2  'fmStyleDropDownList
      TabIndex        =   3
      Top             =   1260
      Width           =   4455
   End
   Begin VB.Label lblRawMaterial
      Caption         =   "Raw_Material property"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblDescription
      Caption         =   "Description property"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblPartNumber
      Caption         =   "Part number / filename"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblSource
      Caption         =   "Workbook / sheet"
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   6615
   End
   Begin VB.Label lblInstructions
      Caption         =   "Choose the spreadsheet column for each SolidWorks property. Optional properties can be left as Do not update."
      Height          =   435
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "ModifiedPartsListMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mCancelled As Boolean
Private mFirstColumn As Long
Private mPartNumberColumn As Long
Private mDescriptionColumn As Long
Private mRawMaterialColumn As Long

Public Sub Configure(ByVal sheet As Object, ByVal headerRow As Long, _
                     ByVal workbookName As String, ByVal sheetName As String)
    mCancelled = True
    mFirstColumn = sheet.UsedRange.Column
    Dim lastColumn As Long
    lastColumn = mFirstColumn + sheet.UsedRange.Columns.Count - 1

    lblSource.Caption = workbookName & "  >  " & sheetName & _
                        "  (header row " & headerRow & ")"
    cboDescription.AddItem "(Do not update)"
    cboRawMaterial.AddItem "(Do not update)"

    Dim columnIndex As Long
    Dim label As String
    For columnIndex = mFirstColumn To lastColumn
        label = ColumnLetter(columnIndex) & " | " & _
                Trim$(CStr(sheet.Cells(headerRow, columnIndex).Value2))
        cboPartNumber.AddItem label
        cboDescription.AddItem label
        cboRawMaterial.AddItem label
    Next columnIndex

    SelectBestMatch cboPartNumber, sheet, headerRow, _
                    Array("part number", "part no", "part #", "filename", "file name", "item")
    SelectBestMatch cboDescription, sheet, headerRow, _
                    Array("description", "part description", "part desc")
    SelectBestMatch cboRawMaterial, sheet, headerRow, _
                    Array("raw_material", "raw material", "material", "material spec")
    If cboPartNumber.ListIndex < 0 And cboPartNumber.ListCount > 0 Then
        cboPartNumber.ListIndex = 0
    End If
    If cboDescription.ListIndex < 0 Then cboDescription.ListIndex = 0
    If cboRawMaterial.ListIndex < 0 Then cboRawMaterial.ListIndex = 0
End Sub

Private Sub SelectBestMatch(ByVal combo As Object, ByVal sheet As Object, _
                            ByVal headerRow As Long, ByVal aliases As Variant)
    Dim columnIndex As Long
    Dim aliasIndex As Long
    Dim normalizedHeader As String
    For columnIndex = mFirstColumn To _
            mFirstColumn + sheet.UsedRange.Columns.Count - 1
        normalizedHeader = LCase$(Trim$(CStr( _
            sheet.Cells(headerRow, columnIndex).Value2)))
        For aliasIndex = LBound(aliases) To UBound(aliases)
            If normalizedHeader = LCase$(CStr(aliases(aliasIndex))) Then
                If combo Is cboPartNumber Then
                    combo.ListIndex = columnIndex - mFirstColumn
                Else
                    combo.ListIndex = columnIndex - mFirstColumn + 1
                End If
                Exit Sub
            End If
        Next aliasIndex
    Next columnIndex
End Sub

Private Sub cmdRun_Click()
    If cboPartNumber.ListIndex < 0 Then
        MsgBox "Choose the part-number/filename column.", vbExclamation, _
               "Map Modified Parts List Columns"
        Exit Sub
    End If
    mPartNumberColumn = mFirstColumn + cboPartNumber.ListIndex
    If cboDescription.ListIndex > 0 Then
        mDescriptionColumn = mFirstColumn + cboDescription.ListIndex - 1
    Else
        mDescriptionColumn = 0
    End If
    If cboRawMaterial.ListIndex > 0 Then
        mRawMaterialColumn = mFirstColumn + cboRawMaterial.ListIndex - 1
    Else
        mRawMaterialColumn = 0
    End If
    If mDescriptionColumn = 0 And mRawMaterialColumn = 0 Then
        MsgBox "Map at least one property column.", vbExclamation, _
               "Map Modified Parts List Columns"
        Exit Sub
    End If
    mCancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    mCancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    mCancelled = True
    Cancel = True
    Me.Hide
End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = mCancelled
End Property

Public Property Get PartNumberColumn() As Long
    PartNumberColumn = mPartNumberColumn
End Property

Public Property Get DescriptionColumn() As Long
    DescriptionColumn = mDescriptionColumn
End Property

Public Property Get RawMaterialColumn() As Long
    RawMaterialColumn = mRawMaterialColumn
End Property

Private Function ColumnLetter(ByVal columnNumber As Long) As String
    Do While columnNumber > 0
        columnNumber = columnNumber - 1
        ColumnLetter = Chr$(65 + (columnNumber Mod 26)) & ColumnLetter
        columnNumber = columnNumber \ 26
    Loop
End Function
