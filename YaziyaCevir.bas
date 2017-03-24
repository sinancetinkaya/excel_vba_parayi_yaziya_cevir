Function YaziyaCevir(sayi)
    basamaklar = Array( _
    Array("", "Bir", "İki", "Üç", "Dört", "Beş", "Altı", "Yedi", "Sekiz", "Dokuz"), _
    Array("", "On", "Yirmi", "Otuz", "Kırk", "Elli", "Altmış", "Yetmiş", "Seksen", "Doksan"), _
    Array("", "Yüz", "İkiYüz", "ÜçYüz", "DörtYüz", "BeşYüz", "AltıYüz", "YediYüz", "SekizYüz", "DokuzYüz") _
    )
    suffixes = Array("", "Bin", "Milyon", "Milyar", "Trilyon", "Katrilyon")
    
    sayi = split(CStr(sayi), Application.DecimalSeparator)
    tamsayi = Val(sayi(0))
    kesir = Val(sayi(1))
    result = ""
    gbs = 3 'grup basamak sayısı
    
    For x = 1 To 2
        If x = 1 Then
            sayi = tamsayi
            birim = "TL"
        Else
            sayi = kesir
            birim = "KRŞ"
        End If
        
        If sayi > 0 Then
            my_string = CStr(sayi)
            suffix = 0
            yaziyla = ""
            
            While Len(my_string) > 0
                If Len(my_string) > gbs Then
                    grup = Right(my_string, gbs)
                    my_string = Left(my_string, Len(my_string) - gbs)
                Else
                    grup = my_string
                    my_string = ""
                End If
                
                yazi = ""
                sayilar = parse(grup)
                sayilar_idx = UBound(sayilar)
                For i = 0 To UBound(sayilar)
                    sayi = Val(sayilar(i))
                    yazi = yazi + basamaklar(sayilar_idx)(sayi)
                    sayilar_idx = sayilar_idx - 1
                Next i
                If yazi = "Bir" And suffix = 1 Then yazi = ""  'BirBin hatasını düzelt
                yazi = yazi + suffixes(suffix)
                suffix = suffix + 1
                yaziyla = yazi + yaziyla
            Wend
            result = result + yaziyla + birim
        End If
    Next x
    
    YaziyaCevir = result
End Function


Function parse(my_string)
    ReDim buff(Len(my_string) - 1)
    For idx = 1 To Len(my_string)
        buff(idx - 1) = Mid$(my_string, idx, 1)
    Next
    parse = buff
End Function

