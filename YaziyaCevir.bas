Function YaziyaCevir(sayi)
    basamaklar = Array( _
    Array("", "Bir", "İki", "Üç", "Dört", "Beş", "Altı", "Yedi", "Sekiz", "Dokuz"), _
    Array("", "On", "Yirmi", "Otuz", "Kırk", "Elli", "Altmış", "Yetmiş", "Seksen", "Doksan"), _
    Array("", "Yüz", "İkiYüz", "ÜçYüz", "DörtYüz", "BeşYüz", "AltıYüz", "YediYüz", "SekizYüz", "DokuzYüz") _
    )
    suffixes = Array("", "Bin", "Milyon", "Milyar", "Trilyon", "Katrilyon")
    para_birimi = Array("TL", "KRŞ")
    
    splitted_numbers = split(CStr(sayi), Application.DecimalSeparator)
    If UBound(splitted_numbers) > 0 Then
        ' Eğer kuruş hanesi tek rakam ise sonuna 0 ekle. 1,2 ise 1,20 yap
        If Len(splitted_numbers(1)) = 1 Then splitted_numbers(1) = splitted_numbers(1) + "0"
        ' Eğer kuruş hanesi 2 rakamdan fazla ise sadece ilk 2 sayıyı al
        If Len(splitted_numbers(1)) > 2 Then splitted_numbers(1) = Left(splitted_numbers(1), 2)
    End If
    
    result = ""
    gbs = 3 'grup basamak sayısı
    
    For x = LBound(splitted_numbers) To UBound(splitted_numbers)
        
        If Val(splitted_numbers(x)) > 0 Then
            str_number = splitted_numbers(x)
            suffix = 0
            yaziyla = ""
            
            While Len(str_number) > 0
                If Len(str_number) > gbs Then
                    grup = Right(str_number, gbs)
                    str_number = Left(str_number, Len(str_number) - gbs)
                Else
                    grup = str_number
                    str_number = ""
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
            result = result + yaziyla + para_birimi(x)
        End If
    Next x
    
    YaziyaCevir = result
End Function


Function parse(str_number)
    ReDim buff(Len(str_number) - 1)
    For idx = 1 To Len(str_number)
        buff(idx - 1) = Mid$(str_number, idx, 1)
    Next
    parse = buff
End Function

