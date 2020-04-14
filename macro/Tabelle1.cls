VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    'to do limit range for data type
    
    
    'check if we monitor variables
   If Target.Column = 6 And Range("a1").value = "Monitor on" Then
        
        Dim buffer As Byte 'buffer for value from cell
        
        dataPointer = Target.row - 3
        addrOffset = Modul12.Verbindung.Data(dataPointer).addrOffset
        addrBit = Modul12.Verbindung.Data(dataPointer).addrBit
        area = Modul12.Verbindung.Data(dataPointer).area
        areaNumber = Modul12.Verbindung.Data(dataPointer).areaNumber
        bits = Modul12.Verbindung.Data(dataPointer).bits
        value = ActiveWorkbook.Worksheets("VarTab").Cells(Target.row, 5)
        
'--------------------------------Write bit--------------------------------------
       If bits = 1 Then 'blue
          Call ToggleBit(addrOffset, addrBit, area, areaNumber, Target.row)
          
'--------------------------------Write non-bit-----------------------------------
       ElseIf bits > 1 Then  'Target.Interior.Color <> RGB(51, 204, 255)
              
'--------------------------------Value non numeric--------------------------------------
              If Not IsNumeric(value) Then
                    ActiveWorkbook.Worksheets("VarTab").Cells(Target.row, 5) = ""
                    MsgBox "Write data cell must be numeric value"
              
'--------------------------------Empty cell--------------------------------------
              ElseIf IsEmpty(value) Then
                    MsgBox "Write data cell is empty"
                    
'--------------------------------Write non-bit--------------------------------------
              Else
                  Call WriteNonBit(addrOffset, area, areaNumber, value, bits)
             End If
       End If
       
       Cancel = True 'turn off edit mode
    End If
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Range("a1").value = "Monitor on" Then
        If (Target.Column = 3 Or Target.Column = 4) And Range("a1").value = "Monitor on" Then
            Beep
            MsgBox "You cannot modify this column during monitoring"
            Cells(1, "A").Select
        End If
    End If
End Sub
