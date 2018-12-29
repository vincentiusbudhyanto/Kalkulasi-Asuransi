Function Hitung_MultiLapis(ByVal AkumulasiKlaim As Double, ByVal LapisanManfaat As Variant, Optional ByVal KlaimTerakhir As Double) As Variant
    
    ReDim SisaManfaat(UBound(LapisanManfaat))
    ReDim TagihanManfaat(UBound(LapisanManfaat))
    
    For i = 0 To UBound(LapisanManfaat)
        
        SisaManfaat(i) = LapisanManfaat(i) - AkumulasiKlaim
                        
        Select Case True
            
            Case SisaManfaat(i) > 0
                AkumulasiKlaim = 0
                
            Case SisaManfaat(i) <= 0
                SisaManfaat(i) = 0
                AkumulasiKlaim = AkumulasiKlaim - LapisanManfaat(i)
                                                                                              
        End Select
        
    Next i
    
    If KlaimTerakhir > 0 Then
                
        For i = 0 To UBound(SisaManfaat)
            
            If SisaManfaat(i) > 0 Then
            
                Sisa = SisaManfaat(i) - KlaimTerakhir
                
                Select Case True
                    
                    Case Sisa > 0
                        SisaManfaat(i) = SisaManfaat(i) - KlaimTerakhir
                        KlaimTerakhir = 0
                        TagihanManfaat(i) = KlaimTerakhir
                        GoTo Selesai
                        
                    Case Sisa <= 0
                        TagihanManfaat(i) = SisaManfaat(i)
                        KlaimTerakhir = KlaimTerakhir - SisaManfaat(i)
                        SisaManfaat(i) = 0
                                                                                                      
                End Select
            
            Else: TagihanManfaat(i) = 0
            
            End If
            
        Next i
        
    ElseIf KlaimTerakhir = 0 Then GoTo Selesai
    
    End If
    
Selesai:
    Hitung_MultiLapis = Array(SisaManfaat, TagihanManfaat, AkumulasiKlaim, KlaimTerakhir)

    GoTo Keluar

Galat:
    Catatan = "GALAT>>Hitung_MultiLapis>>" & Err.Description & ">>" & Err.Number
    Debug.Print Catatan: 'Catat Catatan, "Log_Galat"
    
    GoTo Keluar
    
Keluar:
End Function
