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


'bits Interior.Color = RGB(51, 204, 255)
'byte Interior.Color = RGB(255, 136, 55)
'word Interior.Color = RGB(255, 255, 102)
'dword Interior.Color = RGB(255, 102, 255)
'real Interior.Color = RGB(153, 204, 0)
