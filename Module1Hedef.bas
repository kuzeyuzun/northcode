Attribute VB_Name = "Module1"
Sub UpdateOrCreateMart2025Sheet()
    Dim ws As Worksheet
    Dim shtName As String
    shtName = "Mart 2025"
    
    ' "Mart 2025" sayfas� var m� kontrol et; varsa kullan, yoksa olu�tur.
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(shtName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = shtName
    End If
    
    ' ***********************************************
    ' 1. AY B�LG�LER�N�N YER ALACA�I ALANI (K2:L4)
    ' ***********************************************
    ws.Range("K2").Value = "G�ncel Ay"
    ' L2: Sistem tarihine g�re "mmmm yyyy" format�nda ay (ilk harf b�y�k)
    ws.Range("L2").Value = Application.WorksheetFunction.Proper(Format(Date, "mmmm yyyy"))
    
    ws.Range("K3").Value = "G�n Say�s�"
    Dim currentYear As Long, currentMonth As Long, dayCount As Long
    currentYear = Year(Date)
    currentMonth = Month(Date)
    dayCount = Day(DateSerial(currentYear, currentMonth + 1, 0))
    ws.Range("L3").Value = dayCount
    
    ws.Range("K4").Value = "Ayl�k Hedef"
    ' L4 h�cresi; kullan�c� taraf�ndan manuel doldurulacak. (Makro L4'� de�i�tirmiyor.)
    
    ' ***********************************************
    ' 2. G�NL�K TABLO ALANININ (A, B, D, F, G, H) G�NCELLENMES�
    ' ***********************************************
    Dim startRow As Long, endRow As Long
    startRow = 3
    endRow = dayCount + 2
    ' Sadece otomatik hesaplanan alanlar temizlensin; E s�tunu (Ger�ekle�en Ciro) silinmesin.
    ws.Range("A" & startRow & ":B" & endRow).ClearContents
    ws.Range("D" & startRow & ":D" & endRow).ClearContents
    ws.Range("F" & startRow & ":F" & endRow).ClearContents
    ws.Range("G" & startRow & ":G" & endRow).ClearContents
    ws.Range("H" & startRow & ":H" & endRow).ClearContents
    ws.Range("C" & startRow & ":C" & endRow).ClearContents
    
    ' G�nl�k tablonun ba�l�k sat�r�n� (A1) g�ncelleyelim:
    ws.Range("A1").Value = ws.Range("L2").Value & " G�nl�k Hedef Da��l�m�"
    
    Dim i As Long, r As Long, monthName As String
    ' L2'den ay ismini al (�rn. "Mart")
    monthName = Split(ws.Range("L2").Value, " ")(0)
    
    For i = 1 To dayCount
        r = i + 2   ' Veriler 3. sat�rdan ba�l�yor.
        
        ' A s�tunu: "1 Mart", "2 Mart" vb.
        ws.Cells(r, 1).Value = i & " " & monthName
        ' B s�tunu: �lgili g�n�n hafta ad� (ilk harfi b�y�k)
        ws.Cells(r, 2).Value = Application.WorksheetFunction.Proper(Format(DateSerial(currentYear, currentMonth, i), "dddd"))
        
        ' C s�tunu bo� b�rak�l�yor.
        
        ' D s�tunu: G�nl�k hedef (planlanan) = Ayl�k Hedef/L3; L4 bo�sa bo� d�nd�r.
        ws.Cells(r, 4).FormulaLocal = "=E�ER($L$4="""" ;"""" ; $L$4/$L$3)"
        
        ' E s�tunu: Ger�ekle�en Ciro; makro buraya dokunmuyor, kullan�c� manuel girecek.
        
        ' F s�tunu: Hedef Ger�ekle�me = E�er E bo�sa bo�, yoksa E/D
        ws.Cells(r, 6).FormulaLocal = "=E�ER(E" & r & "="""" ;"""" ; E" & r & "/D" & r & ")"
        
        ' G s�tunu: Hedeften Sapma = E�er E bo�sa bo�, yoksa F-1
        ws.Cells(r, 7).FormulaLocal = "=E�ER(E" & r & "="""" ;"""" ; F" & r & "-1)"
        
        ' H s�tunu: G�ncel Hedef hesaplama
        ' - �lk g�n (r=3) i�in bo� b�rak.
        ' - Di�er g�nlerde: E�er L4 bo�sa bo�;
        '   E�er bug�ne kadar (E3:E(r-1)) hen�z hi� de�er girilmemi�se, planlanan g�nl�k hedefi (L4/L3) yaz;
        '   Aksi halde, (Ayl�k Hedef - TOPLA(E3:E(r-1))) / (L3 - (SATIR()-3))
        If r = 3 Then
            ws.Cells(r, 8).Value = ""
        Else
            ws.Cells(r, 8).FormulaLocal = "=E�ER($L$4="""" ;"""" ; E�ER(SAY($E$3:E" & (r - 1) & ")=0; $L$4/$L$3; ($L$4-TOPLA($E$3:E" & (r - 1) & "))/($L$3-(SATIR()-3))))"
        End If
    Next i
    
    ' ***********************************************
    ' 3. SAYFA B���MLEND�RMES�
    ' ***********************************************
    ' F ve G s�tunlar� y�zde bi�iminde (ondal�k g�stermeden)
    ws.Range("F" & startRow & ":F" & endRow).NumberFormat = "0%"
    ws.Range("G" & startRow & ":G" & endRow).NumberFormat = "0%"
    
    ' E ve H s�tunlar� binlik ayrac�, ondal�k g�stermeden
    ws.Range("D" & startRow & ":D" & endRow).NumberFormat = "#,##0"
    ws.Range("E" & startRow & ":E" & endRow).NumberFormat = "#,##0"
    ws.Range("H" & startRow & ":H" & endRow).NumberFormat = "#,##0"
    
    MsgBox "Mart 2025 sayfas�, manuel girilen veriler korunarak g�ncellendi.", vbInformation
End Sub

