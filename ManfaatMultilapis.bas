Const PemisahUmum = "$!Sfe4O$#"

Function Hitung_AkumulasiKlaim(ByVal NoPeserta As String, ByVal daftarNoPeserta As Variant, ByVal daftarNoKlaim As Variant, ByVal daftarKlaimDisetujui As Variant, _
    Optional ByVal Identifikator As String, Optional ByVal daftarIdentifikator As Variant, Optional ByVal KodeKeluarga As String, Optional ByVal daftarNoKeluarga As Variant)

    If Identifikator <> "" Then
    
        For i = 0 To UBound(daftarNoPeserta)
        
            If daftarNoPeserta(i) = NoPeserta Then
            
                If LokData = "" Then LokData = i Else LokData = LokData & PemisahUmum & i
                        
            End If
        
        Next i
        
        LokData = Split(LokData, PemisahUmum)
    
        For i = 0 To UBound(LokData)
        
            If daftarIdentifikator(LokData(i)) = Identifikator Then
            
                If x2LokData = "" Then x2LokData = i Else x2LokData = x2LokData & PemisahUmum & i
                        
            End If
        
        Next i
        
        x2LokData = Split(x2LokData, PemisahUmum)
        
        GoTo LoncatanIdentifikator
        
    End If

    For i = 0 To UBound(daftarNoPeserta)
    
        If daftarNoPeserta(i) = NoPeserta Then
        
            If x2LokData = "" Then x2LokData = i Else x2LokData = x2LokData & PemisahUmum & i
                    
        End If
    
    Next i
    
    x2LokData = Split(x2LokData, PemisahUmum)
    
LoncatanIdentifikator:
    
    ReDim AkumulasiKlaim(UBound(x2LokData))
    For i = 0 To UBound(x2LokData)
    
        NoKlaim = daftarNoKlaim(x2LokData(i))
        
        For j = 0 To UBound(x2LokData)
        
            If daftarNoKlaim(x2LokData(j)) < NoKlaim Then
                
                AkumulasiKlaim(i) = daftarKlaimDisetujui(j) + AkumulasiKlaim(i)
                
            End If
        
        Next j
        
    Next i
    
    If KodeKeluarga <> "" Then
        
        
        
    End If
    
    Hitung_AkumulasiKlaim = Array(x2LokData, AkumulasiKlaim)

    GoTo Keluar

Galat:

    GoTo Keluar

Keluar:

End Function

Function Hitung_MultiLapis(ByVal AkumulasiKlaim As Double, ByVal LapisanManfaat As Variant, Optional ByVal KlaimTerakhir As Double) As Variant
    
    On Error GoTo Galat
    
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
    
    Hitung_MultiLapis = False
    
    GoTo Keluar
    
Keluar:
End Function

Private Sub Uji_Hitung_MultiLapis()

    AkumulasiKlaim = 10000000
    LapisanManfaat = Array(9000000, 2000000, 5000000, 2000000)
    KlaimTerakhir = 6000000
        
    Hitung = Hitung_MultiLapis(AkumulasiKlaim, LapisanManfaat, KlaimTerakhir)
    
End Sub
