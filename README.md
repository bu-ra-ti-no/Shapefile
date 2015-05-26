# ESRI-Shape-file-library
Very simple library for working with shape files.

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

F = New ShapeFile(SpatialCategories.Point)
F.CodePage = CodePages.Russian_Windows
F.Fields.Add("Поле1", Field.FieldTypes.String)

F.AddNew()

ReDim F.Geometry.Boundaries(0)
F.Geometry.Boundaries(0) = New Boundary
ReDim F.Geometry.Boundaries(0).Points(0)
F.Geometry.Boundaries(0).Points(0) = New Point(10, 20)
F.Record.Item(0) = "Точечный объект"

F.Save("Новый.shp")
```
