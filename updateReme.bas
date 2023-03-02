Attribute VB_Name = "updateReme"
Option Compare Database


' Update pol w tabeli [Remeasure]
' Wyzerowanie station(value) w tabeli [POSTPLOT] 
' przeniesienie danych z tabeli [REMEASURE]->[POSTPLOT]

' przygotowal DJurkowski@geofizyka.pl


 Sub UpdateX()

        Dim dbs As Database
        Dim qdf As QueryDef

        ' Zmien sciezke do bazy danych jesli potrzeba

        Set dbs = OpenDatabase("C:\CZ-2\database\CZ-2.mdb")

        ' ZAKTUALIZUJ zakresy lini SP i RP jesli potrzeba (pierwsze SP, drugie RP),
        ' Zaktualizuj nazwy pol jesli masz inne

        dbs.Execute "UPDATE [REMEASURE],[POSTPLOT]" _
& "SET [REMEASURE].`COG Local Easting` = [POSTPLOT].`COG Local Easting`," _
& "[REMEASURE].`COG Local Northing` = [POSTPLOT].`COG Local Northing`," _
& "[REMEASURE].`COG Local Height` = [POSTPLOT].`COG Local Height`," _
& "[REMEASURE].`Acquired_Julian_Day` = [POSTPLOT].`Acquired_Julian_Day`," _
& "[REMEASURE].`Descriptor` = [POSTPLOT].`Descriptor`," _
& "[REMEASURE].`Comment` = [POSTPLOT].`Comment`," _
& "[REMEASURE].`Indeks` = [POSTPLOT].`Indeks`," _
& "[REMEASURE].`Description1` = [POSTPLOT].`Description1`," _
& "[REMEASURE].`Description2` = [POSTPLOT].`Description2`," _
& "[REMEASURE].`OfficeNote` = [POSTPLOT].`OfficeNote`," _
& "[REMEASURE].`PPV` = [POSTPLOT].`PPV`," _
& "[REMEASURE].`Status` = iif([POSTPLOT].`Track` Between 3001 And 3469, '2', '1')" _
& "WHERE" _
& "[POSTPLOT].`Station (value)` > 0 and datediff ('d',[REMEASURE].`Survey Time (Local)`,Now()) = 0 And" _
& "([POSTPLOT].`Track` Between 3001 And 3469 or [POSTPLOT].`Track` Between 1998 And 2484)  And" _
& "[REMEASURE].`Station (value)` = [POSTPLOT].`Station (value)` and [REMEASURE].`Indeks` = [POSTPLOT].`Indeks`;"


        dbs.Execute "UPDATE [REMEASURE],[POSTPLOT]" _
& "SET [POSTPLOT].`Station (value)` = '0'" _
& "WHERE" _
& "[POSTPLOT].`Station (value)` > 0 And datediff ('d',[REMEASURE].`Survey Time (Local)`,Now()) = 0 And" _
& "([POSTPLOT].`Track` Between 3001 And 3469 or [POSTPLOT].`Track` Between 1998 And 2484)  And" _
& "[REMEASURE].`Station (value)` = [POSTPLOT].`Station (value)`and [REMEASURE].`Indeks` = [POSTPLOT].`Indeks`;"

        dbs.Execute "Insert into [POSTPLOT]" _
        & "SELECT * FROM [REMEASURE] WHERE datediff ('d',[REMEASURE].`Survey Time (Local)`,Now()) = 0;"
        
        dbs.Close
     
    End Sub


