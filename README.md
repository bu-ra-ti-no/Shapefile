# ESRI Shapefile library

![GitHub repo size](https://img.shields.io/github/repo-size/bu-ra-ti-no/Shapefile?style=plastic)
![GitHub top language](https://img.shields.io/github/languages/top/bu-ra-ti-no/Shapefile?style=plastic)

Very simple VB.NET library for working with ESRI shapefiles.

Supports reading and writing of all types of shapefiles.
Does not support the memo fields.

```vb.net
'Библиотека для работы с шейп-файлами.

'============================================

'Работа с существующим файлом:

Dim F As ShapeFile, S As String

If Not IO.File.Exists("Водохранилища.shp") Then Stop
F = ShapeFile.Load("Водохранилища.shp")

'F.Bookmark = 0
S = "Общее количество геометрий: " & F.RecordCount & Chr(13) & Chr(13)
S += "Количество границ в первой геометрии: " & F.Geometry.Boundaries.Length & Chr(13)
S += "Количество вершин в первой геометрии: " & F.Geometry.PointCount & Chr(13)
S += "Значение в поле " & F.Fields(0).ColumnName & ": " & F.Record(0).ToString & Chr(13) & Chr(13)
F.Bookmark = 1
S += "Количество границ во второй геометрии: " & F.Geometry.Boundaries.Length & Chr(13)
S += "Количество вершин во второй геометрии: " & F.Geometry.PointCount & Chr(13)
S += "Значение в поле " & F.Fields(0).ColumnName & ": " & F.Record(0).ToString

MsgBox(S)

'============================================

'Создание нового файла:

F = New ShapeFile(SpatialCategories.PolyLine)
F.CodePage = CodePages.Russian_Windows
F.Fields.Add("Поле1", Field.FieldTypes.String)

F.AddNew()

F.Geometry.Boundaries = New Boundary() {New Boundary}
F.Geometry.Boundaries(0).Points = New Point() {New Point(10, 20), New Point(20, 20)}
F.Record.Item(0) = "Линейный объект"

F.Save("Новый.shp")
```
