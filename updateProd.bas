Attribute VB_Name = "updateProd"
Option Compare Database

' Update pol w postplocie na podstawie danych z tabeli [AVG]
' przygotowal DJurkowski@geofizyka.pl

 Sub UpdateProd()

        Dim dbs As Database
        Dim qdf As QueryDef

        ' Zmien sciezke do bazy danych jesli potrzeba

        Set dbs = OpenDatabase("C:\CZ-2\database\CZ-2.mdb")


        ' ZAKTUALIZUJ zakresy lini SP, update dot. daty, update dot. statusu

        dbs.Execute "UPDATE [AVG],[POSTPLOT]" _
& "SET [POSTPLOT].`COG Local Easting` = [AVG].`Local Easting`," _
& "[POSTPLOT].`COG Local Northing` = [AVG].`Local Northing`," _
& "[POSTPLOT].`COG Local Height` = [AVG].`Height`," _
& "[POSTPLOT].`Acquired_Julian_Day` = IIF([AVG].`Julian Day`<100, '20230'&[AVG].`Julian Day`, '2023'&[AVG].`Julian Day`)," _
& "[POSTPLOT].`Status` = IIF([AVG].`Descriptor` like 4, 3, IIF(SQR(([POSTPLOT].`Local Easting`-[AVG].`Local Easting`)^2 + ([POSTPLOT].`Local Northing`-[AVG].`Local Northing`)^2)<5.20, 4, 5))" _
& "WHERE" _
& "[POSTPLOT].`Station (value)` > 0 And" _
& "[POSTPLOT].`Track` Between 3001 And 3469  And" _
& "[AVG].`Station (value)` = [POSTPLOT].`Station (value)`;"

        dbs.Close
     
    End Sub
