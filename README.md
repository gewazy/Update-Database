# Update-Database

Skrypty do MS Access

## updateReme.bas
- aktualizyje wszystkie istotne pola w tabeli REMEASURE 
- zeruje `station (value)` w tabeli [POSTPLOT]
- kopiuje dane z tabeli [REMEASURE] do tabeli [POSTPLOT]

## updateProd.bas
Aktualizuje dane w tabeli [POSTPLOT] na podstawie danych z tabeli [AVG]
- COG Local Easting
- COG Local Northing
- COG Local height
- Acquired_Julian_Date
- Status
