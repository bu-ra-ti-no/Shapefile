Option Strict On
Option Infer On

Imports System.IO
Imports System.Globalization

''' <summary>
''' Класс для работы с шейп-файлами.
''' </summary>
<System.ComponentModel.DesignerCategory("Code")> _
Public Class ShapeFile

    Public Enum CodePages
        [Default] = 0
        US_MSDOS = 437
        International_MSDOS = 850
        Windows_ANSI = 1252
        Greek_MSDOS = 737
        EasernEuropean_MSDOS = 852
        Chinese_Windows_950 = 950
        Chinese_Windows_936 = 936
        Japanese_Windows = 932
        Russian_MSDOS = 866
        Russian_Windows = 1251
        Greek_Windows = 1253
        Turkish_Windows = 1254
        Eastern_European_Windows = 1250
        Arabic_Windows = 1256
        Hebrew_Windows = 1255
    End Enum

    ''' <summary>
    ''' Code page for new instances.
    ''' </summary>
    Public Shared DefaultCodePage As CodePages

    ''' <summary>
    ''' Code page.
    ''' </summary>
    Public CodePage As CodePages

    ''' <summary>
    ''' WKT projection for new instances.
    ''' </summary>
    ''' <example>
    ''' GEOGCS["GCS_WGS_1984",DATUM["D_WGS_1984",SPHEROID["WGS_1984",6378137.0,298.257223563]],PRIMEM["Greenwich",0.0],UNIT["Degree",0.0174532925199433]]
    ''' </example>
    Public Shared DefaultProjectionWKT As String = ""

    ''' <summary>
    ''' WKT projection string.
    ''' </summary>
    ''' <example>
    ''' GEOGCS["GCS_WGS_1984",DATUM["D_WGS_1984",SPHEROID["WGS_1984",6378137.0,298.257223563]],PRIMEM["Greenwich",0.0],UNIT["Degree",0.0174532925199433]]
    ''' </example>
    ''' <remarks>
    ''' If this parameter is not specified, then .prj-file will not be created.
    ''' </remarks>
    Public ProjectionWKT As String = ""

    Public Enum SpatialCategories
        NullShape = 0
        Point = 1
        PolyLine = 3
        Polygon = 5
        MultiPoint = 8
        PointZ = 11
        PolyLineZ = 13
        PolygonZ = 15
        MultiPointZ = 18
        PointM = 21
        PolyLineM = 23
        PolygonM = 25
        MultiPointM = 28
        MultiPatch = 31
    End Enum

    Private _SpatialCategory As SpatialCategories
    ''' <summary>
    ''' Геометрическая характеристика объектов.
    ''' </summary>
    Public ReadOnly Property SpatialCategory() As SpatialCategories
        Get
            Return _SpatialCategory
        End Get
    End Property

    Public Shared Function SpatialCategoryFromFile(ByVal FileName As String) As SpatialCategories
        Dim FileNameSHP As String
        Dim Ret As SpatialCategories
        FileNameSHP = RemoveExtension(FileName) & ".shp"
        Using sp = New FileStream(FileNameSHP, FileMode.Open, FileAccess.Read, FileShare.Read)
            sp.Seek(32, SeekOrigin.Begin)
            Ret = CType(sp.ReadByte, SpatialCategories)
            sp.Close()
        End Using
        Return Ret
    End Function

    Public Sub New(ByVal SpatialCategory As SpatialCategories)
        _SpatialCategory = SpatialCategory
        ProjectionWKT = DefaultProjectionWKT
        CodePage = DefaultCodePage
    End Sub
    Private Sub New()
        CodePage = DefaultCodePage
        ProjectionWKT = DefaultProjectionWKT
    End Sub

    ''' <summary>
    ''' Создаёт новый экземпляр, инициализированный указанным файлом.
    ''' </summary>
    ''' <param name="FileName">Путь к файлу.</param>
    Public Shared Function Load(ByVal FileName As String) As ShapeFile
        If FileName.Length = 0 Then Return Nothing
        Dim FileNameSHP, FileNameSHX, FileNameDBF, FileNamePRJ As String

        FileNameSHP = RemoveExtension(FileName) & ".shp"
        FileNameSHX = RemoveExtension(FileName) & ".shx"
        FileNameDBF = RemoveExtension(FileName) & ".dbf"
        FileNamePRJ = RemoveExtension(FileName) & ".prj"

        Dim Ret As New ShapeFile

        If Not File.Exists(FileNameSHP) Then Return Nothing
        If Not File.Exists(FileNameSHX) Then
            If Not CreateSHX(FileNameSHP) Then Return Nothing
        End If

        If File.Exists(FileNameDBF) Then
            Using sf = New FileStream(FileNameDBF, FileMode.Open, FileAccess.Read, FileShare.Read)
                If Not Ret.LoadRecords(sf) Then Return Nothing
                sf.Close()
            End Using
        Else
            Ret.Table = Nothing
        End If

        Using _
        sp = New FileStream(FileNameSHP, FileMode.Open, FileAccess.Read, FileShare.Read), _
        sx = New FileStream(FileNameSHX, FileMode.Open, FileAccess.Read, FileShare.Read)
            Dim sprj As FileStream = Nothing
            If File.Exists(FileNamePRJ) Then
                sprj = New FileStream(FileNamePRJ, FileMode.Open, FileAccess.Read, FileShare.Read)
            End If
            If Not Ret.LoadGeometries(sp, sx, sprj) Then Return Nothing
            sp.Close()
            sx.Close()
            If sprj IsNot Nothing Then sprj.Close()
        End Using

        Return Ret
    End Function

    ''' <summary>
    ''' Сохранение shape-файла.
    ''' </summary>
    ''' <param name="FileName">Имя файла.</param>
    Public Function Save(ByVal FileName As String) As Boolean
        If FileName.Length = 0 Then Return False
        Dim FileNameSHP, FileNameSHX, FileNameDBF, FileNamePRJ As String

        FileNameSHP = RemoveExtension(FileName) & ".shp"
        FileNameSHX = RemoveExtension(FileName) & ".shx"
        FileNameDBF = RemoveExtension(FileName) & ".dbf"
        FileNamePRJ = RemoveExtension(FileName) & ".prj"

        Using _
        sp = New FileStream(FileNameSHP, FileMode.Create, FileAccess.Write), _
        sx = New FileStream(FileNameSHX, FileMode.Create, FileAccess.Write)
            Dim sprj As FileStream = Nothing
            If ProjectionWKT.Length > 0 Then
                sprj = New FileStream(FileNamePRJ, FileMode.Create, FileAccess.Write)
            End If
            If Not SaveGeometries(sp, sx, sprj) Then Return False
            sp.Close()
            sx.Close()
            If sprj IsNot Nothing Then sprj.Close()
        End Using

        If Table IsNot Nothing Then
            Using sf = New FileStream(FileNameDBF, FileMode.Create, FileAccess.Write)
                If Not SaveRecords(sf) Then Return False
            End Using
        End If

        Return True
    End Function

    Private _Bookmark As Integer
    ''' <summary>
    ''' Указатель на текущую геометрию.
    ''' </summary>
    Public Property Bookmark() As Integer
        Get
            Return _Bookmark
        End Get
        Set(ByVal value As Integer)
            If value < 0 Then Exit Property
            If value > _RecordCount - 1 Then Exit Property
            _Bookmark = value
            _Geometry = Geometries(value)
            If Table IsNot Nothing Then
                _Row = DirectCast(Table.Rows(value), Record)
            End If
        End Set
    End Property

    ''' <summary>
    ''' Создаёт новую геометрию и добавляет её в конец набора. 
    ''' Перемещает указатель на добавленную геометрию.
    ''' </summary>
    Public Function AddNew() As Integer
        _Geometry = New Geometry
        _Geometry.SpatialCategory = _SpatialCategory
        Geometries.Add(_Geometry)
        If Table IsNot Nothing Then
            _Row = Table.NewRow()
            Table.Rows.Add(_Row)
        End If
        _Bookmark = _RecordCount
        _RecordCount += 1
        Return _Bookmark
    End Function

    ''' <summary>
    ''' Удаляет текущую геометрию.
    ''' </summary>
    ''' <remarks>
    ''' Если удаляемая геометрия находится в конце набора, 
    ''' указатель <see cref="Bookmark"/> уменьшается на единицу.
    ''' </remarks>
    Public Sub Delete()
        If Geometries.Count > 0 Then
            Geometries.RemoveAt(_Bookmark)
            If Table IsNot Nothing Then
                Table.Rows.RemoveAt(_Bookmark)
            End If
            _RecordCount -= 1
            If _RecordCount = 0 Then
                _Geometry = Nothing
                _Row = Nothing
            Else
                If _Bookmark > _RecordCount - 1 Then
                    _Bookmark = _RecordCount - 1
                End If
                Bookmark = _Bookmark
            End If
        End If
    End Sub

    Private _RecordCount As Integer
    Private ReadOnly Geometries As New List(Of Geometry)
    Private Table As New Table
    Private _Fields As New FieldCollection(Table.Columns)
    Private _Geometry As Geometry
    Private _Row As Record

    ''' <summary>
    ''' Текущая геометрия.
    ''' </summary>
    Public ReadOnly Property Geometry() As Geometry
        Get
            Return _Geometry
        End Get
    End Property

    ''' <summary>
    ''' Количество геометрий в наборе.
    ''' </summary>
    ''' <remarks>
    ''' Количество записей всегда равно количеству геометрий.
    ''' </remarks>
    Public ReadOnly Property RecordCount() As Integer
        Get
            Return _RecordCount
        End Get
    End Property

    ''' <summary>
    ''' Счисывает счётчик геометрий из указанного шейп-файла.
    ''' </summary>
    Public Shared Function RecordCountFromFile(ByVal FileName As String) As Integer
        Dim FileNameSHP As String
        Dim Ret As Integer
        FileNameSHP = RemoveExtension(FileName) & ".shp"
        Using sp = New FileStream(FileNameSHP, FileMode.Open, FileAccess.Read, FileShare.Read)
            sp.Seek(24, SeekOrigin.Begin)
            Ret = (BitConverterEx.GetInt32Big(sp) - 50) \ 4
            sp.Close()
        End Using
        Return Ret
    End Function

    ''' <summary>
    ''' Набор полей семантической таблицы.
    ''' </summary>
    Public ReadOnly Property Fields() As FieldCollection
        Get
            If Table Is Nothing Then Return Nothing
            Return _Fields
        End Get
    End Property

    ''' <summary>
    ''' Семантическая информация, связанная с текущей геометрией.
    ''' </summary>
    ''' <remarks>
    ''' Если DBF-файл отсутствует, возвращается пустая ссылка.
    ''' </remarks>
    Public ReadOnly Property Record() As Record
        Get
            If Table Is Nothing Then Return Nothing
            Return DirectCast(Table.Rows(_Bookmark), Record)
        End Get
    End Property

    ''' <summary>
    ''' Вычисляет минимальный ограничивающий прямоугольник, занимаемый 
    ''' геометриями.
    ''' </summary>
    ''' <remarks>
    ''' Если нет ни одной геометрии, границы прямоугольника растянуты 
    ''' до значений <c>Double.MaxValue</c> и инвертированы.
    ''' </remarks>
    Public ReadOnly Property MBB() As MBB
        Get
            Dim Box As MBB
            Box.Reset()
            For Each G In Geometries
                Box.AddMBB(G.MBB)
            Next
            Return Box
        End Get
    End Property

    ''' <summary>
    ''' Создаёт индексный файл для <paramref name="FileNameSHP"/>.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> в случае успешного завершения операции, 
    ''' <c>False</c> в противном случае.
    ''' </returns>
    ''' <remarks>
    ''' Индексный файл можно построить только в том случае, когда
    ''' шейп-файл имеет стандартныме смещения.
    ''' </remarks>
    Private Shared Function CreateSHX(ByVal FileNameSHP As String) As Boolean
        Dim FileNameSHX = RemoveExtension(FileNameSHP) & ".shx"
        Dim b(99) As Byte
        Using sp As New FileStream(FileNameSHP, FileMode.Open, FileAccess.Read, FileShare.Read), _
              sx As New FileStream(FileNameSHX, FileMode.Create, FileAccess.Write)
            sp.Read(b, 0, 100)
            sx.Write(b, 0, 100)
            Dim MaxPos = BitConverterEx.GetInt32Big(b, 24)
            MaxPos += MaxPos - 1
            Dim Pos = 104, Shift = 50, rc = 0
            Do While Pos < MaxPos
                rc += 1
                sp.Seek(Pos, SeekOrigin.Begin)
                Dim L = BitConverterEx.GetInt32Big(sp)
                BitConverterEx.SetBytesBig(sx, Shift)
                BitConverterEx.SetBytesBig(sx, L)
                Pos += 8 + L + L
                Shift += 4 + L
            Loop
            sx.Seek(24, SeekOrigin.Begin)
            BitConverterEx.SetBytesBig(sx, 50 + rc * 4)
            sp.Close()
            sx.Close()
        End Using
    End Function

    Private Shared Function RemoveExtension(ByVal FileName As String) As String
        Return FileName.Remove(FileName.Length - Path.GetExtension(FileName).Length)
    End Function

    Private Function LoadRecords(ByVal s As FileStream) As Boolean
        Dim e As System.Text.Encoding
        Dim b(3999) As Byte 'Buffer
        Dim Pos = 1
        Dim F As Field
        If Table Is Nothing Then
            Table = New Table
        Else
            Table.Clear()
            Table.Columns.Clear()
        End If
        '=======================================================================
        'Code page
        '=======================================================================
        s.Seek(29, SeekOrigin.Begin)
        Select Case s.ReadByte
            Case &H1 'US_MSDOS
                CodePage = CodePages.US_MSDOS
            Case &H2 'International_MSDOS
                CodePage = CodePages.International_MSDOS
            Case &H3 'Windows_ANSI
                CodePage = CodePages.Windows_ANSI
            Case &H6A 'Greek_MSDOS
                CodePage = CodePages.Greek_MSDOS
            Case &H64 'EasernEuropean_MSDOS
                CodePage = CodePages.EasernEuropean_MSDOS
            Case &H78 'Chinese_Windows
                CodePage = CodePages.Chinese_Windows_950
            Case &H7A 'Chinese_Windows
                CodePage = CodePages.Chinese_Windows_936
            Case &H7B 'Japanese_Windows
                CodePage = CodePages.Japanese_Windows
            Case &H26, &H65 'Russian_MSDOS
                CodePage = CodePages.Russian_MSDOS
            Case &HC9, &H57 'Russian_Windows
                CodePage = CodePages.Russian_Windows
            Case &HCB 'Greek_Windows
                CodePage = CodePages.Greek_Windows
            Case &HCA 'Turkish_Windows
                CodePage = CodePages.Turkish_Windows
            Case &HC8 'Eastern_European_Windows
                CodePage = CodePages.Eastern_European_Windows
            Case &H7E 'Arabic_Windows
                CodePage = CodePages.Arabic_Windows
            Case &H7D 'Hebrew_Windows
                CodePage = CodePages.Hebrew_Windows
            Case Else
                CodePage = CodePages.Default
        End Select
        If CodePage = CodePages.Default Then
            e = System.Text.Encoding.Default
        Else
            e = System.Text.Encoding.GetEncoding(CodePage)
        End If
        '=======================================================================
        'Header
        '=======================================================================
        s.Seek(4, SeekOrigin.Begin)
        _RecordCount = BitConverterEx.GetInt32Little(s)
        Dim HeaderSize = BitConverterEx.GetInt16Little(s)
        Dim RecordSize = BitConverterEx.GetInt16Little(s)
        Dim FieldCount = (HeaderSize - 1) \ 32 - 1
        '=======================================================================
        'Fields
        '=======================================================================
        s.Seek(32, SeekOrigin.Begin)
        For i = 0 To FieldCount - 1
            s.Read(b, 0, 32)
            Dim j = 10
            Do While b(j) = 0
                j -= 1
            Loop
            Dim FieldName = e.GetString(b, 0, j + 1)
            Dim FieldType As Field.FieldTypes
            Select Case b(11)
                Case 67, 99 'Asc("C")
                    FieldType = Field.FieldTypes.String
                Case 78, 110 'Asc("N")
                    FieldType = Field.FieldTypes.Numeric
                Case 76, 108 'Asc("L")
                    FieldType = Field.FieldTypes.Logical
                Case 68, 100 'Asc("D")
                    FieldType = Field.FieldTypes.Date
                Case 70, 102 'Asc("F")
                    FieldType = Field.FieldTypes.Float
                Case Else  '77,109 Asc("M")
                    FieldType = Field.FieldTypes.Memo
            End Select
            F = New Field(FieldName, FieldType, Len:=b(16), Dec:=b(17))
            F.Pos = Pos
            Pos += F.Len
            Table.Columns.Add(F)
        Next
        If s.ReadByte() <> &HD Then Return False
        '=======================================================================
        'Records
        '=======================================================================
        Dim p = New CultureInfo("en-us")
        For i = 0 To _RecordCount - 1
            Dim Row = Table.NewRow()
            s.Read(b, 0, RecordSize)
            For Each F In Table.Columns
                Select Case F.FieldType
                    Case Field.FieldTypes.String
                        Dim v = e.GetString(b, F.Pos, F.Len).TrimEnd
                        If v.Length = 0 Then
                            Row.Item(F) = DBNull.Value
                        Else
                            Row.Item(F) = v
                        End If
                    Case Field.FieldTypes.Numeric
                        If b(F.Pos + F.Len - 1) = &H20 Then
                            Row.Item(F) = DBNull.Value
                        ElseIf F.Dec = 0 Then
                            Row.Item(F) = Int32.Parse(e.GetString(b, F.Pos, F.Len).TrimStart)
                        Else
                            Row.Item(F) = Double.Parse(e.GetString(b, F.Pos, F.Len).TrimStart, p)
                        End If
                    Case Field.FieldTypes.Float
                        If b(F.Pos + F.Len - 1) = &H20 Then
                            Row.Item(F) = DBNull.Value
                        Else
                            Row.Item(F) = Double.Parse(e.GetString(b, F.Pos, F.Len).TrimStart, p)
                        End If
                    Case Field.FieldTypes.Date
                        If b(F.Pos) = &H20 Then
                            Row.Item(F) = DBNull.Value
                        Else
                            Row.Item(F) = Date.ParseExact(e.GetString(b, F.Pos, F.Len), "yyyyMMdd", CultureInfo.InvariantCulture)
                        End If
                    Case Field.FieldTypes.Logical
                        Select Case b(F.Pos)
                            Case 84, 116 'Asc("T")
                                Row.Item(F) = True
                            Case 70, 102 'Asc("F")
                                Row.Item(F) = False
                            Case Else
                                Row.Item(F) = DBNull.Value
                        End Select
                    Case Else
                        '?
                End Select
            Next
            Table.Rows.Add(Row)
        Next
        Return True
    End Function

    Private Function LoadGeometries(ByVal sp As FileStream, ByVal sx As FileStream, ByVal sprj As FileStream) As Boolean

        If sprj IsNot Nothing Then
            Dim b(CInt(sprj.Length) - 1) As Byte
            sprj.Read(b, 0, b.Length)
            ProjectionWKT = Text.Encoding.ASCII.GetString(b)
        End If

        Dim Buffer() As Byte
        sp.Seek(32, SeekOrigin.Begin)
        _SpatialCategory = CType(BitConverterEx.GetInt32Little(sp), SpatialCategories)
        sx.Seek(24, SeekOrigin.Begin)
        _RecordCount = (BitConverterEx.GetInt32Big(sx) - 50) \ 4
        sx.Seek(100, SeekOrigin.Begin)
        For i = 1 To _RecordCount
            Dim G = New Geometry
            Dim Shift = BitConverterEx.GetInt32Big(sx) * 2 + 8
            Dim L = BitConverterEx.GetInt32Big(sx) * 2
            ReDim Buffer(L - 1)
            sp.Seek(Shift, SeekOrigin.Begin)
            sp.Read(Buffer, 0, L)
            G.SpatialCategory = CType(Buffer(0), SpatialCategories)
            Select Case G.SpatialCategory
                Case SpatialCategories.NullShape
                Case SpatialCategories.Point
                    G.CreateBoundaries(1)
                    G.Boundaries(0).CreatePoints(1)
                    G.Boundaries(0).Points(0).X = BitConverterEx.GetDoubleLittle(Buffer, 4)
                    G.Boundaries(0).Points(0).Y = BitConverterEx.GetDoubleLittle(Buffer, 12)
                Case SpatialCategories.MultiPoint
                    G.CreateBoundaries(1)
                    Dim NumPoints = BitConverterEx.GetInt32Little(Buffer, 36)
                    G.Boundaries(0).CreatePoints(NumPoints)
                    For j = 0 To NumPoints - 1
                        G.Boundaries(0).Points(j).X = BitConverterEx.GetDoubleLittle(Buffer, 40 + j * 16)
                        G.Boundaries(0).Points(j).Y = BitConverterEx.GetDoubleLittle(Buffer, 48 + j * 16)
                    Next
                Case SpatialCategories.PolyLine, SpatialCategories.Polygon
                    Dim NumParts = BitConverterEx.GetInt32Little(Buffer, 36)
                    Dim NumPoints = BitConverterEx.GetInt32Little(Buffer, 40)
                    G.CreateBoundaries(NumParts)
                    Dim Origin = 44 + 4 * NumParts
                    Dim Points(NumPoints - 1) As Point
                    For j = 0 To NumPoints - 1
                        Points(j) = New Point
                        Points(j).X = BitConverterEx.GetDoubleLittle(Buffer, Origin + j * 16)
                        Points(j).Y = BitConverterEx.GetDoubleLittle(Buffer, Origin + j * 16 + 8)
                    Next
                    Dim Count, Pos, Pos1 As Integer
                    Pos1 = NumPoints
                    For j = NumParts - 1 To 0 Step -1
                        Pos = BitConverterEx.GetInt32Little(Buffer, 44 + 4 * j)
                        Count = Pos1 - Pos
                        Pos1 = Pos
                        ReDim G.Boundaries(j).Points(Count - 1)
                        Array.Copy(Points, Pos, G.Boundaries(j).Points, 0, Count)
                    Next
                Case SpatialCategories.PointM
                    G.CreateBoundaries(1)
                    G.Boundaries(0).CreatePoints(1)
                    G.Boundaries(0).Points(0).X = BitConverterEx.GetDoubleLittle(Buffer, 4)
                    G.Boundaries(0).Points(0).Y = BitConverterEx.GetDoubleLittle(Buffer, 12)
                    G.Boundaries(0).Points(0).M = BitConverterEx.GetDoubleLittle(Buffer, 20)
                Case SpatialCategories.MultiPointM
                    G.CreateBoundaries(1)
                    Dim NumPoints = BitConverterEx.GetInt32Little(Buffer, 36)
                    G.Boundaries(0).CreatePoints(NumPoints)
                    For j = 0 To NumPoints - 1
                        G.Boundaries(0).Points(j).X = BitConverterEx.GetDoubleLittle(Buffer, 40 + j * 16)
                        G.Boundaries(0).Points(j).Y = BitConverterEx.GetDoubleLittle(Buffer, 48 + j * 16)
                        G.Boundaries(0).Points(j).M = BitConverterEx.GetDoubleLittle(Buffer, 56 + 16 * NumPoints + j * 8)
                    Next
                Case SpatialCategories.PolyLineM, SpatialCategories.PolygonM
                    Dim NumParts = BitConverterEx.GetInt32Little(Buffer, 36)
                    Dim NumPoints = BitConverterEx.GetInt32Little(Buffer, 40)
                    G.CreateBoundaries(NumParts)
                    Dim Origin = 44 + 4 * NumParts
                    Dim Points(NumPoints - 1) As Point
                    For j = 0 To NumPoints - 1
                        Points(j) = New Point
                        Points(j).X = BitConverterEx.GetDoubleLittle(Buffer, Origin + j * 16)
                        Points(j).Y = BitConverterEx.GetDoubleLittle(Buffer, Origin + j * 16 + 8)
                        Points(j).M = BitConverterEx.GetDoubleLittle(Buffer, Origin + NumPoints * 16 + 16 + j * 8)
                    Next
                    Dim Count, Pos, Pos1 As Integer
                    Pos1 = NumPoints
                    For j = NumParts - 1 To 0 Step -1
                        Pos = BitConverterEx.GetInt32Little(Buffer, 44 + 4 * j)
                        Count = Pos1 - Pos
                        Pos1 = Pos
                        ReDim G.Boundaries(j).Points(Count - 1)
                        Array.Copy(Points, Pos, G.Boundaries(j).Points, 0, Count)
                    Next
                Case SpatialCategories.PointZ
                    G.CreateBoundaries(1)
                    G.Boundaries(0).CreatePoints(1)
                    G.Boundaries(0).Points(0).X = BitConverterEx.GetDoubleLittle(Buffer, 4)
                    G.Boundaries(0).Points(0).Y = BitConverterEx.GetDoubleLittle(Buffer, 12)
                    G.Boundaries(0).Points(0).Z = BitConverterEx.GetDoubleLittle(Buffer, 20)
                    G.Boundaries(0).Points(0).M = BitConverterEx.GetDoubleLittle(Buffer, 28)
                Case SpatialCategories.MultiPointZ
                    G.CreateBoundaries(1)
                    Dim NumPoints = BitConverterEx.GetInt32Little(Buffer, 36)
                    G.Boundaries(0).CreatePoints(NumPoints)
                    For j = 0 To NumPoints - 1
                        G.Boundaries(0).Points(j).X = BitConverterEx.GetDoubleLittle(Buffer, 40 + j * 16)
                        G.Boundaries(0).Points(j).Y = BitConverterEx.GetDoubleLittle(Buffer, 48 + j * 16)
                        G.Boundaries(0).Points(j).Z = BitConverterEx.GetDoubleLittle(Buffer, 56 + 16 * NumPoints + j * 8)
                        G.Boundaries(0).Points(j).M = BitConverterEx.GetDoubleLittle(Buffer, 72 + 24 * NumPoints + j * 8)
                    Next
                Case SpatialCategories.PolyLineZ, SpatialCategories.PolygonZ
                    Dim NumParts = BitConverterEx.GetInt32Little(Buffer, 36)
                    Dim NumPoints = BitConverterEx.GetInt32Little(Buffer, 40)
                    G.CreateBoundaries(NumParts)
                    Dim Origin = 44 + 4 * NumParts
                    Dim Points(NumPoints - 1) As Point
                    For j = 0 To NumPoints - 1
                        Points(j) = New Point
                        Points(j).X = BitConverterEx.GetDoubleLittle(Buffer, Origin + j * 16)
                        Points(j).Y = BitConverterEx.GetDoubleLittle(Buffer, Origin + j * 16 + 8)
                        Points(j).Z = BitConverterEx.GetDoubleLittle(Buffer, Origin + NumPoints * 16 + 16 + j * 8)
                        Points(j).M = BitConverterEx.GetDoubleLittle(Buffer, Origin + NumPoints * 24 + 32 + j * 8)
                    Next
                    Dim Count, Pos, Pos1 As Integer
                    Pos1 = NumPoints
                    For j = NumParts - 1 To 0 Step -1
                        Pos = BitConverterEx.GetInt32Little(Buffer, 44 + 4 * j)
                        Count = Pos1 - Pos
                        Pos1 = Pos
                        ReDim G.Boundaries(j).Points(Count - 1)
                        Array.Copy(Points, Pos, G.Boundaries(j).Points, 0, Count)
                    Next
                Case SpatialCategories.MultiPatch
                    Dim NumParts = BitConverterEx.GetInt32Little(Buffer, 36)
                    Dim NumPoints = BitConverterEx.GetInt32Little(Buffer, 40)
                    G.CreateBoundaries(NumParts)
                    Dim Origin = 44 + 4 * NumParts
                    For j = 0 To NumParts - 1
                        G.Boundaries(j).PartType = CType(BitConverterEx.GetInt32Little(Buffer, Origin), PartTypes)
                        Origin += 4
                    Next
                    Dim Points(NumPoints - 1) As Point
                    For j = 0 To NumPoints - 1
                        Points(j) = New Point
                        Points(j).X = BitConverterEx.GetDoubleLittle(Buffer, Origin + j * 16)
                        Points(j).Y = BitConverterEx.GetDoubleLittle(Buffer, Origin + j * 16 + 8)
                        Points(j).Z = BitConverterEx.GetDoubleLittle(Buffer, Origin + NumPoints * 16 + 16 + j * 8)
                        Points(j).M = BitConverterEx.GetDoubleLittle(Buffer, Origin + NumPoints * 24 + 32 + j * 8)
                    Next
                    Dim Count, Pos, Pos1 As Integer
                    Pos1 = NumPoints
                    For j = NumParts - 1 To 0 Step -1
                        Pos = BitConverterEx.GetInt32Little(Buffer, 44 + 4 * j)
                        Count = Pos1 - Pos
                        Pos1 = Pos
                        ReDim G.Boundaries(j).Points(Count - 1)
                        Array.Copy(Points, Pos, G.Boundaries(j).Points, 0, Count)
                    Next
            End Select
            Geometries.Add(G)
        Next
        Bookmark = 0
        Return True
    End Function

    Private Function SaveRecords(ByVal s As FileStream) As Boolean
        Dim b(31) As Byte
        '=======================================================================
        'File Header
        '=======================================================================
        b(0) = 3
        With Date.Today
            b(1) = CByte(.Year - 1900)
            b(2) = CByte(.Year - 1900)
            b(3) = CByte(.Day)
        End With
        BitConverterEx.SetBytesLittle(b, 4, _RecordCount)
        BitConverterEx.SetBytesLittle(b, 8, CShort(32 * (Table.Columns.Count + 1) + 1))
        Dim L = 1S
        For Each F As Field In Table.Columns
            L = CShort(L + F.Len)
        Next
        BitConverterEx.SetBytesLittle(b, 10, L)
        Select Case CodePage
            Case CodePages.US_MSDOS
                b(29) = &H1
            Case CodePages.International_MSDOS
                b(29) = &H2
            Case CodePages.Windows_ANSI
                b(29) = &H3
            Case CodePages.Greek_MSDOS
                b(29) = &H6A
            Case CodePages.EasernEuropean_MSDOS
                b(29) = &H64
            Case CodePages.Chinese_Windows_950
                b(29) = &H78
            Case CodePages.Chinese_Windows_936
                b(29) = &H7A
            Case CodePages.Japanese_Windows
                b(29) = &H7B
            Case CodePages.Russian_MSDOS
                b(29) = &H26
            Case CodePages.Russian_Windows
                b(29) = &HC9
            Case CodePages.Greek_Windows
                b(29) = &HCB
            Case CodePages.Turkish_Windows
                b(29) = &HCA
            Case CodePages.Eastern_European_Windows
                b(29) = &HC8
            Case CodePages.Arabic_Windows
                b(29) = &H7E
            Case CodePages.Hebrew_Windows
                b(29) = &H7D
        End Select
        s.Write(b, 0, 32)
        '=======================================================================
        'Field Headers
        '=======================================================================
        Dim e As System.Text.Encoding
        If CodePage = CodePages.Default Then
            e = System.Text.Encoding.Default
        Else
            e = System.Text.Encoding.GetEncoding(CodePage)
        End If
        Dim offset = 1
        For Each F As Field In Table.Columns
            Array.Clear(b, 0, 32)
            Dim n = e.GetBytes(F.ColumnName)
            Array.Copy(n, b, Math.Min(n.Length, 11))
            Select Case F.FieldType
                Case Field.FieldTypes.Numeric
                    b(11) = 78
                Case Field.FieldTypes.Float
                    b(11) = 70
                Case Field.FieldTypes.String
                    b(11) = 67
                Case Field.FieldTypes.Date
                    b(11) = 70
                Case Field.FieldTypes.Logical
                    b(11) = 76
                Case Else
            End Select
            BitConverterEx.SetBytesLittle(b, 12, offset)
            b(16) = CByte(F.Len)
            b(17) = CByte(F.Dec)
            s.Write(b, 0, 32)
            offset += F.Len
        Next
        s.WriteByte(&HD)
        '=======================================================================
        'Records
        '=======================================================================
        Dim p = New CultureInfo("en-us")
        Dim r(L - 1) As Byte, Empty(L - 1) As Byte
        For i = 0 To L - 1
            Empty(i) = &H20
        Next
        For i = 0 To _RecordCount - 1
            Array.Copy(Empty, r, L)
            Dim Row = DirectCast(Table.Rows(i), Record)
            Dim Pos = 1
            For Each F As Field In Table.Columns
                Select Case F.FieldType
                    Case Field.FieldTypes.Numeric
                        If Row.Item(F) IsNot DBNull.Value Then
                            If F.Dec = 0 Then
                                Dim v = e.GetBytes(CInt(Row.Item(F)).ToString)
                                Array.Copy(v, 0, r, Pos + F.Len - v.Length, v.Length)
                            Else
                                Dim v = e.GetBytes(CDbl(Row.Item(F)).ToString(p))
                                Array.Copy(v, 0, r, Pos + F.Len - v.Length, v.Length)
                            End If
                        End If
                    Case Field.FieldTypes.Float
                        If Row.Item(F) IsNot DBNull.Value Then
                            Dim v = e.GetBytes(CDbl(Row.Item(F)).ToString(p))
                            Array.Copy(v, 0, r, Pos + F.Len - v.Length, v.Length)
                        End If
                    Case Field.FieldTypes.String
                        If Row.Item(F) IsNot DBNull.Value Then
                            Dim v = e.GetBytes(CStr(Row.Item(F)))
                            Array.Copy(v, 0, r, Pos, Math.Min(v.Length, 254))
                        End If
                    Case Field.FieldTypes.Date
                        If Row.Item(F) IsNot DBNull.Value Then
                            Dim D = CDate(Row.Item(F))
                            Array.Copy(e.GetBytes(D.ToString("yyyyMMdd")), 0, r, Pos, 8)
                        End If
                    Case Field.FieldTypes.Logical
                        If Row.Item(F) IsNot DBNull.Value Then
                            If CBool(Row.Item(F)) Then
                                r(Pos) = 84
                            Else
                                r(Pos) = 70
                            End If
                        End If
                    Case Else
                End Select
                Pos += F.Len
            Next
            s.Write(r, 0, r.Length)
        Next
        s.WriteByte(&H1A) 'EOF
        s.Flush()
    End Function

    Private Function SaveGeometries(ByVal sp As FileStream, ByVal sx As FileStream, ByVal sprj As FileStream) As Boolean
        '=======================================================================
        'Projection
        '=======================================================================
        If sprj IsNot Nothing Then
            Dim b = Text.Encoding.ASCII.GetBytes(ProjectionWKT)
            sprj.Seek(0, SeekOrigin.Begin)
            sprj.Write(b, 0, b.Length)
        End If
        '=======================================================================
        'Main File Records
        '=======================================================================
        sp.Seek(100, SeekOrigin.Begin)
        sx.Seek(100, SeekOrigin.Begin)
        Dim L As Integer, G As Geometry
        For i = 0 To _RecordCount - 1
            '===================================================================
            'Record Header
            '===================================================================
            Dim Shift = CInt(sp.Position)
            BitConverterEx.SetBytesBig(sp, i + 1)
            G = Geometries(i)
            Select Case G.SpatialCategory
                Case SpatialCategories.NullShape
                    L = 4
                Case SpatialCategories.Point
                    L = 20
                Case SpatialCategories.MultiPoint
                    L = 40 + G.PointCount(0) * 16
                Case SpatialCategories.PolyLine, SpatialCategories.Polygon
                    L = 44 + G.Boundaries.Length * 4 + G.PointCount(-1) * 16
                Case SpatialCategories.PointM
                    L = 28
                Case SpatialCategories.MultiPointM
                    L = 56 + G.PointCount(0) * 24
                Case SpatialCategories.PolyLineM, SpatialCategories.PolygonM
                    L = 56 + G.Boundaries.Length * 4 + G.PointCount(-1) * 24
                Case SpatialCategories.PointZ
                    L = 36
                Case SpatialCategories.MultiPointZ
                    L = 72 + G.PointCount(0) * 32
                Case SpatialCategories.PolyLineZ, SpatialCategories.PolygonZ
                    L = 76 + G.Boundaries.Length * 4 + G.PointCount(-1) * 32
                Case SpatialCategories.MultiPatch
                    L = 76 + G.Boundaries.Length * 8 + G.PointCount(-1) * 32
            End Select
            BitConverterEx.SetBytesBig(sp, L >> 1)
            '===================================================================
            'Record Contents
            '===================================================================
            Dim Buffer(L - 1) As Byte
            BitConverterEx.SetBytesLittle(Buffer, 0, CInt(G.SpatialCategory))
            Select Case G.SpatialCategory
                Case SpatialCategories.NullShape
                Case SpatialCategories.Point
                    With G.Boundaries(0).Points(0)
                        BitConverterEx.SetBytesLittle(Buffer, 4, .X)
                        BitConverterEx.SetBytesLittle(Buffer, 12, .Y)
                    End With
                Case SpatialCategories.MultiPoint
                    With G.MBB
                        BitConverterEx.SetBytesLittle(Buffer, 4, .x1)
                        BitConverterEx.SetBytesLittle(Buffer, 12, .y1)
                        BitConverterEx.SetBytesLittle(Buffer, 20, .x2)
                        BitConverterEx.SetBytesLittle(Buffer, 28, .y2)
                    End With
                    BitConverterEx.SetBytesLittle(Buffer, 36, G.PointCount(0))
                    Dim Pos = 40
                    For Each P In G.Boundaries(0).Points
                        BitConverterEx.SetBytesLittle(Buffer, Pos, P.X)
                        Pos += 8
                        BitConverterEx.SetBytesLittle(Buffer, Pos, P.Y)
                        Pos += 8
                    Next
                Case SpatialCategories.PolyLine, SpatialCategories.Polygon
                    With G.MBB
                        BitConverterEx.SetBytesLittle(Buffer, 4, .x1)
                        BitConverterEx.SetBytesLittle(Buffer, 12, .y1)
                        BitConverterEx.SetBytesLittle(Buffer, 20, .x2)
                        BitConverterEx.SetBytesLittle(Buffer, 28, .y2)
                    End With
                    BitConverterEx.SetBytesLittle(Buffer, 36, G.Boundaries.Length)
                    Dim cnt, Pos As Integer
                    cnt = 0 : Pos = 44
                    For j = 0 To G.Boundaries.Length - 1
                        BitConverterEx.SetBytesLittle(Buffer, Pos, cnt)
                        Pos += 4
                        cnt += G.Boundaries(j).Points.Length
                    Next
                    BitConverterEx.SetBytesLittle(Buffer, 40, cnt)
                    For j = 0 To G.Boundaries.Length - 1
                        For Each P In G.Boundaries(j).Points
                            BitConverterEx.SetBytesLittle(Buffer, Pos, P.X)
                            Pos += 8
                            BitConverterEx.SetBytesLittle(Buffer, Pos, P.Y)
                            Pos += 8
                        Next
                    Next
                Case SpatialCategories.PointM
                    With G.Boundaries(0).Points(0)
                        BitConverterEx.SetBytesLittle(Buffer, 4, .X)
                        BitConverterEx.SetBytesLittle(Buffer, 12, .Y)
                        BitConverterEx.SetBytesLittle(Buffer, 20, .M)
                    End With
                Case SpatialCategories.MultiPointM
                    Dim MBB = G.MBB
                    BitConverterEx.SetBytesLittle(Buffer, 4, MBB.x1)
                    BitConverterEx.SetBytesLittle(Buffer, 12, MBB.y1)
                    BitConverterEx.SetBytesLittle(Buffer, 20, MBB.x2)
                    BitConverterEx.SetBytesLittle(Buffer, 28, MBB.y2)
                    BitConverterEx.SetBytesLittle(Buffer, 36, G.PointCount(0))
                    Dim Pos = 40
                    For Each P In G.Boundaries(0).Points
                        BitConverterEx.SetBytesLittle(Buffer, Pos, P.X)
                        Pos += 8
                        BitConverterEx.SetBytesLittle(Buffer, Pos, P.Y)
                        Pos += 8
                    Next
                    'M
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.m1)
                    Pos += 8
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.m2)
                    Pos += 8
                    For Each P In G.Boundaries(0).Points
                        BitConverterEx.SetBytesLittle(Buffer, Pos, P.M)
                        Pos += 8
                    Next
                Case SpatialCategories.PolyLineM, SpatialCategories.PolygonM
                    Dim MBB = G.MBB
                    BitConverterEx.SetBytesLittle(Buffer, 4, MBB.x1)
                    BitConverterEx.SetBytesLittle(Buffer, 12, MBB.y1)
                    BitConverterEx.SetBytesLittle(Buffer, 20, MBB.x2)
                    BitConverterEx.SetBytesLittle(Buffer, 28, MBB.y2)
                    BitConverterEx.SetBytesLittle(Buffer, 36, G.Boundaries.Length)
                    Dim cnt, Pos As Integer
                    Pos = 44
                    For j = 0 To G.Boundaries.Length - 1
                        BitConverterEx.SetBytesLittle(Buffer, Pos, cnt)
                        Pos += 4
                        cnt += G.Boundaries(j).Points.Length
                    Next
                    BitConverterEx.SetBytesLittle(Buffer, 40, cnt)
                    For j = 0 To G.Boundaries.Length - 1
                        For Each P In G.Boundaries(j).Points
                            BitConverterEx.SetBytesLittle(Buffer, Pos, P.X)
                            Pos += 8
                            BitConverterEx.SetBytesLittle(Buffer, Pos, P.Y)
                            Pos += 8
                        Next
                    Next
                    'M
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.m1)
                    Pos += 8
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.m2)
                    Pos += 8
                    For j = 0 To G.Boundaries.Length - 1
                        For Each P In G.Boundaries(j).Points
                            BitConverterEx.SetBytesLittle(Buffer, Pos, P.M)
                            Pos += 8
                        Next
                    Next
                Case SpatialCategories.PointZ
                    With G.Boundaries(0).Points(0)
                        BitConverterEx.SetBytesLittle(Buffer, 4, .X)
                        BitConverterEx.SetBytesLittle(Buffer, 12, .Y)
                        BitConverterEx.SetBytesLittle(Buffer, 20, .Z)
                        BitConverterEx.SetBytesLittle(Buffer, 28, .M)
                    End With
                Case SpatialCategories.MultiPointZ
                    Dim MBB = G.MBB
                    BitConverterEx.SetBytesLittle(Buffer, 4, MBB.x1)
                    BitConverterEx.SetBytesLittle(Buffer, 12, MBB.y1)
                    BitConverterEx.SetBytesLittle(Buffer, 20, MBB.x2)
                    BitConverterEx.SetBytesLittle(Buffer, 28, MBB.y2)
                    BitConverterEx.SetBytesLittle(Buffer, 36, G.PointCount(0))
                    Dim Pos = 40
                    For Each P In G.Boundaries(0).Points
                        BitConverterEx.SetBytesLittle(Buffer, Pos, P.X)
                        Pos += 8
                        BitConverterEx.SetBytesLittle(Buffer, Pos, P.Y)
                        Pos += 8
                    Next
                    'Z
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.z1)
                    Pos += 8
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.z2)
                    Pos += 8
                    For Each P In G.Boundaries(0).Points
                        BitConverterEx.SetBytesLittle(Buffer, Pos, P.Z)
                        Pos += 8
                    Next
                    'M
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.m1)
                    Pos += 8
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.m2)
                    Pos += 8
                    For Each P In G.Boundaries(0).Points
                        BitConverterEx.SetBytesLittle(Buffer, Pos, P.M)
                        Pos += 8
                    Next
                Case SpatialCategories.PolyLineZ, SpatialCategories.PolygonZ
                    Dim MBB = G.MBB
                    BitConverterEx.SetBytesLittle(Buffer, 4, MBB.x1)
                    BitConverterEx.SetBytesLittle(Buffer, 12, MBB.y1)
                    BitConverterEx.SetBytesLittle(Buffer, 20, MBB.x2)
                    BitConverterEx.SetBytesLittle(Buffer, 28, MBB.y2)
                    BitConverterEx.SetBytesLittle(Buffer, 36, G.Boundaries.Length)
                    Dim cnt, Pos As Integer
                    cnt = 0 : Pos = 44
                    For j = 0 To G.Boundaries.Length - 1
                        BitConverterEx.SetBytesLittle(Buffer, Pos, cnt)
                        Pos += 4
                        cnt += G.Boundaries(j).Points.Length
                    Next
                    BitConverterEx.SetBytesLittle(Buffer, 40, cnt)
                    For j = 0 To G.Boundaries.Length - 1
                        For Each P In G.Boundaries(j).Points
                            BitConverterEx.SetBytesLittle(Buffer, Pos, P.X)
                            Pos += 8
                            BitConverterEx.SetBytesLittle(Buffer, Pos, P.Y)
                            Pos += 8
                        Next
                    Next
                    'Z
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.z1)
                    Pos += 8
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.z2)
                    Pos += 8
                    For j = 0 To G.Boundaries.Length - 1
                        For Each P In G.Boundaries(j).Points
                            BitConverterEx.SetBytesLittle(Buffer, Pos, P.Z)
                            Pos += 8
                        Next
                    Next
                    'M
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.m1)
                    Pos += 8
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.m2)
                    Pos += 8
                    For j = 0 To G.Boundaries.Length - 1
                        For Each P In G.Boundaries(j).Points
                            BitConverterEx.SetBytesLittle(Buffer, Pos, P.M)
                            Pos += 8
                        Next
                    Next
                Case SpatialCategories.MultiPatch
                    Dim MBB = G.MBB
                    BitConverterEx.SetBytesLittle(Buffer, 4, MBB.x1)
                    BitConverterEx.SetBytesLittle(Buffer, 12, MBB.y1)
                    BitConverterEx.SetBytesLittle(Buffer, 20, MBB.x2)
                    BitConverterEx.SetBytesLittle(Buffer, 28, MBB.y2)
                    BitConverterEx.SetBytesLittle(Buffer, 36, G.Boundaries.Length)
                    Dim cnt, Pos As Integer
                    cnt = 0 : Pos = 44
                    For j = 0 To G.Boundaries.Length - 1
                        BitConverterEx.SetBytesLittle(Buffer, Pos, cnt)
                        Pos += 4
                        cnt += G.Boundaries(j).Points.Length
                    Next
                    BitConverterEx.SetBytesLittle(Buffer, 40, cnt)
                    For j = 0 To G.Boundaries.Length - 1
                        BitConverterEx.SetBytesLittle(Buffer, Pos, G.Boundaries(j).PartType)
                        Pos += 4
                    Next
                    For j = 0 To G.Boundaries.Length - 1
                        For Each P In G.Boundaries(j).Points
                            BitConverterEx.SetBytesLittle(Buffer, Pos, P.X)
                            Pos += 8
                            BitConverterEx.SetBytesLittle(Buffer, Pos, P.Y)
                            Pos += 8
                        Next
                    Next
                    'Z
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.z1)
                    Pos += 8
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.z2)
                    Pos += 8
                    For j = 0 To G.Boundaries.Length - 1
                        For Each P In G.Boundaries(j).Points
                            BitConverterEx.SetBytesLittle(Buffer, Pos, P.Z)
                            Pos += 8
                        Next
                    Next
                    'M
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.m1)
                    Pos += 8
                    BitConverterEx.SetBytesLittle(Buffer, Pos, MBB.m2)
                    Pos += 8
                    For j = 0 To G.Boundaries.Length - 1
                        For Each P In G.Boundaries(j).Points
                            BitConverterEx.SetBytesLittle(Buffer, Pos, P.M)
                            Pos += 8
                        Next
                    Next
            End Select
            sp.Write(Buffer, 0, L)
            '===================================================================
            'Index File Record
            '===================================================================
            BitConverterEx.SetBytesBig(sx, Shift \ 2)
            BitConverterEx.SetBytesBig(sx, L \ 2)
        Next
        '=======================================================================
        'Main File Header
        '=======================================================================
        Dim Header(99) As Byte
        BitConverterEx.SetBytesBig(Header, 0, 9994)
        BitConverterEx.SetBytesBig(Header, 24, CInt(sp.Position) \ 2)
        BitConverterEx.SetBytesLittle(Header, 28, 1000)
        Header(32) = CByte(_SpatialCategory)
        With MBB
            BitConverterEx.SetBytesLittle(Header, 36, .x1)
            BitConverterEx.SetBytesLittle(Header, 44, .y1)
            BitConverterEx.SetBytesLittle(Header, 52, .x2)
            BitConverterEx.SetBytesLittle(Header, 60, .y2)
            BitConverterEx.SetBytesLittle(Header, 68, .z1)
            BitConverterEx.SetBytesLittle(Header, 76, .z2)
            BitConverterEx.SetBytesLittle(Header, 84, .m1)
            BitConverterEx.SetBytesLittle(Header, 92, .m2)
        End With
        sp.Seek(0, SeekOrigin.Begin)
        sp.Write(Header, 0, 100)
        '=======================================================================
        'Index File Header
        '=======================================================================
        BitConverterEx.SetBytesBig(Header, 24, CInt(sx.Position) \ 2)
        sx.Seek(0, SeekOrigin.Begin)
        sx.Write(Header, 0, 100)
        '=======================================================================
        sp.Flush()
        sx.Flush()
        If sprj IsNot Nothing Then sprj.Flush()
        Return True
    End Function

End Class

''' <summary>
''' Минимальный ограничивающий прямоугольник, занимаемый геометрией.
''' </summary>
Public Structure MBB
    Public x1, y1, x2, y2, z1, z2, m1, m2 As Double
    ''' <summary>
    ''' Инициализация структуры.
    ''' </summary>
    ''' <remarks>
    ''' Метод растягивает границы прямоугольника до <c>Double.MaxValue</c> 
    ''' и инвертирует их.
    ''' </remarks>
    Public Sub Reset()
        x1 = Double.MaxValue
        y1 = Double.MaxValue
        z1 = Double.MaxValue
        m1 = Double.MaxValue
        x2 = Double.MinValue
        y2 = Double.MinValue
        z2 = Double.MinValue
        m2 = Double.MinValue
    End Sub
    ''' <summary>
    ''' Наращивает габариты прямоугольника, чтобы вместить вершину 
    ''' геометрии.
    ''' </summary>
    Public Sub AddPoint(ByVal Point As Point)
        If x1 > Point.X Then x1 = Point.X
        If x2 < Point.X Then x2 = Point.X
        If y1 > Point.Y Then y1 = Point.Y
        If y2 < Point.Y Then y2 = Point.Y
        If z1 > Point.Z Then z1 = Point.Z
        If z2 < Point.Z Then z2 = Point.Z
        If m1 > Point.M Then m1 = Point.M
        If m2 < Point.M Then m2 = Point.M
    End Sub
    ''' <summary>
    ''' Наращивает габариты прямоугольника, чтобы вместить точечные 
    ''' геометрии.
    ''' </summary>
    Public Sub AddPoints(ByVal Points() As Point)
        For Each P In Points
            AddPoint(P)
        Next
    End Sub
    ''' <summary>
    ''' Наращивает габариты прямоугольника, чтобы вместить новый 
    ''' прямоугольник.
    ''' </summary>
    Public Sub AddMBB(ByVal MBB As MBB)
        If x1 > MBB.x1 Then x1 = MBB.x1
        If x2 < MBB.x2 Then x2 = MBB.x2
        If y1 > MBB.y1 Then y1 = MBB.y1
        If y2 < MBB.y2 Then y2 = MBB.y2
        If z1 > MBB.z1 Then z1 = MBB.z1
        If z2 < MBB.z2 Then z2 = MBB.z2
        If m1 > MBB.m1 Then m1 = MBB.m1
        If m2 < MBB.m2 Then m2 = MBB.m2
    End Sub
End Structure

''' <summary>
''' Вершина геометрии.
''' </summary>
Public Class Point
    Public X, Y, Z, M As Double

    Public Sub New()

    End Sub
    Public Sub New(ByVal X As Double, ByVal Y As Double)
        Me.X = X
        Me.Y = Y
    End Sub
    Public Sub New(ByVal X As Double, ByVal Y As Double, ByVal Z As Double)
        Me.X = X
        Me.Y = Y
        Me.Z = Z
    End Sub
    Public Sub New(ByVal X As Double, ByVal Y As Double, ByVal Z As Double, ByVal M As Double)
        Me.X = X
        Me.Y = Y
        Me.Z = Z
        Me.M = M
    End Sub
End Class

''' <summary>
''' Граница геометрии (набор вершин).
''' </summary>
''' <remarks></remarks>
Public Class Boundary
    Public Points() As Point
    Friend Sub CreatePoints(ByVal Count As Integer)
        ReDim Points(Count - 1)
        For i = 0 To Count - 1
            Points(i) = New Point
        Next
    End Sub
    Public PartType As PartTypes = PartTypes.Ring
End Class

Public Enum PartTypes
    TriangleStrip
    TriangleFan
    OuterRing
    InnerRing
    FirstRing
    Ring
End Enum

''' <summary>
''' Геометрия.
''' </summary>
Public Class Geometry
    Public Boundaries() As Boundary
    ''' <summary>
    ''' Количество вершин в границе <paramref name="Boundary"/>.
    ''' </summary>
    ''' <remarks>
    ''' Если номер границы не указан, вернёт общее количество вершин во 
    ''' всех границах.
    ''' </remarks>
    Public ReadOnly Property PointCount(Optional ByVal Boundary As Integer = -1) As Integer
        Get
            If Boundaries Is Nothing Then Return 0
            If Boundary < 0 Then
                Dim cnt As Integer
                For i = 0 To Boundaries.Length - 1
                    cnt += Boundaries(i).Points.Length
                Next
                Return cnt
            ElseIf Boundary <= Boundaries.Length - 1 Then
                Return Boundaries(Boundary).Points.Length
            Else
                Return 0
            End If
        End Get
    End Property
    Public SpatialCategory As ShapeFile.SpatialCategories
    Friend Sub CreateBoundaries(ByVal Count As Integer)
        ReDim Boundaries(Count - 1)
        For i = 0 To Count - 1
            Boundaries(i) = New Boundary
        Next
    End Sub
    ''' <summary>
    ''' Минимальный ограничивающий прямоугольник.
    ''' </summary>
    Public ReadOnly Property MBB() As MBB
        Get
            Dim Box As MBB
            Try
                Box.Reset()
                For Each B In Boundaries
                    Box.AddPoints(B.Points)
                Next
            Catch ex As Exception
            End Try
            Return Box
        End Get
    End Property
End Class

Friend NotInheritable Class BitConverterEx
    Public Shared Function GetInt32Little(ByVal b() As Byte, ByVal StartIndex As Integer) As Int32
        If BitConverter.IsLittleEndian Then
            Return BitConverter.ToInt32(b, StartIndex)
        Else
            Dim a() As Byte = {b(StartIndex + 3), b(StartIndex + 2), b(StartIndex + 1), b(StartIndex)}
            Return BitConverter.ToInt32(a, 0)
        End If
    End Function
    Public Shared Function GetInt32Big(ByVal b() As Byte, ByVal StartIndex As Integer) As Int32
        If BitConverter.IsLittleEndian Then
            Dim a() As Byte = {b(StartIndex + 3), b(StartIndex + 2), b(StartIndex + 1), b(StartIndex)}
            Return BitConverter.ToInt32(a, 0)
        Else
            Return BitConverter.ToInt32(b, StartIndex)
        End If
    End Function
    Public Shared Function GetInt16Little(ByVal b() As Byte, ByVal StartIndex As Integer) As Int16
        If BitConverter.IsLittleEndian Then
            Return BitConverter.ToInt16(b, StartIndex)
        Else
            Dim a() As Byte = {b(StartIndex + 1), b(StartIndex)}
            Return BitConverter.ToInt16(a, 0)
        End If
    End Function
    Public Shared Function GetInt16Big(ByVal b() As Byte, ByVal StartIndex As Integer) As Int16
        If BitConverter.IsLittleEndian Then
            Dim a() As Byte = {b(StartIndex + 1), b(StartIndex)}
            Return BitConverter.ToInt16(a, 0)
        Else
            Return BitConverter.ToInt16(b, StartIndex)
        End If
    End Function
    Public Shared Function GetDoubleLittle(ByVal b() As Byte, ByVal StartIndex As Integer) As Double
        If BitConverter.IsLittleEndian Then
            Return BitConverter.ToDouble(b, StartIndex)
        Else
            Dim a() As Byte = {b(StartIndex + 7), b(StartIndex + 6), b(StartIndex + 5), b(StartIndex + 4), _
                               b(StartIndex + 3), b(StartIndex + 2), b(StartIndex + 1), b(StartIndex)}
            Return BitConverter.ToDouble(a, 0)
        End If
    End Function
    Public Shared Function GetDoubleBig(ByVal b() As Byte, ByVal StartIndex As Integer) As Double
        If BitConverter.IsLittleEndian Then
            Dim a() As Byte = {b(StartIndex + 7), b(StartIndex + 6), b(StartIndex + 5), b(StartIndex + 4), _
                               b(StartIndex + 3), b(StartIndex + 2), b(StartIndex + 1), b(StartIndex)}
            Return BitConverter.ToDouble(a, 0)
        Else
            Return BitConverter.ToDouble(b, StartIndex)
        End If
    End Function
    Public Shared Function GetSingleLittle(ByVal b() As Byte, ByVal StartIndex As Integer) As Single
        If BitConverter.IsLittleEndian Then
            Return BitConverter.ToSingle(b, StartIndex)
        Else
            Dim a() As Byte = {b(StartIndex + 3), b(StartIndex + 2), b(StartIndex + 1), b(StartIndex)}
            Return BitConverter.ToSingle(a, 0)
        End If
    End Function
    Public Shared Function GetSingleBig(ByVal b() As Byte, ByVal StartIndex As Integer) As Single
        If BitConverter.IsLittleEndian Then
            Dim a() As Byte = {b(StartIndex + 3), b(StartIndex + 2), b(StartIndex + 1), b(StartIndex)}
            Return BitConverter.ToSingle(a, 0)
        Else
            Return BitConverter.ToSingle(b, StartIndex)
        End If
    End Function

    Public Shared Function GetInt32Little(ByVal s As Stream) As Int32
        Dim b(3) As Byte
        s.Read(b, 0, 4)
        Return GetInt32Little(b, 0)
    End Function
    Public Shared Function GetInt32Big(ByVal s As Stream) As Int32
        Dim b(3) As Byte
        s.Read(b, 0, 4)
        Return GetInt32Big(b, 0)
    End Function
    Public Shared Function GetInt16Little(ByVal s As Stream) As Int16
        Dim b(1) As Byte
        s.Read(b, 0, 2)
        Return GetInt16Little(b, 0)
    End Function
    Public Shared Function GetInt16Big(ByVal s As Stream) As Int16
        Dim b(1) As Byte
        s.Read(b, 0, 2)
        Return GetInt16Big(b, 0)
    End Function
    Public Shared Function GetDoubleLittle(ByVal s As Stream) As Double
        Dim b(7) As Byte
        s.Read(b, 0, 8)
        Return GetDoubleLittle(b, 0)
    End Function
    Public Shared Function GetDoubleBig(ByVal s As Stream) As Double
        Dim b(7) As Byte
        s.Read(b, 0, 8)
        Return GetDoubleBig(b, 0)
    End Function
    Public Shared Function GetSingleLittle(ByVal s As Stream) As Single
        Dim b(3) As Byte
        s.Read(b, 0, 4)
        Return GetSingleLittle(b, 0)
    End Function
    Public Shared Function GetSingleBig(ByVal s As Stream) As Single
        Dim b(3) As Byte
        s.Read(b, 0, 4)
        Return GetSingleBig(b, 0)
    End Function

    Public Shared Sub SetBytesLittle(ByVal s As Stream, ByVal value As Int16)
        Dim b = GetBytesLittle(value)
        s.Write(b, 0, 2)
    End Sub
    Public Shared Sub SetBytesLittle(ByVal s As Stream, ByVal value As Int32)
        Dim b = GetBytesLittle(value)
        s.Write(b, 0, 4)
    End Sub
    Public Shared Sub SetBytesLittle(ByVal s As Stream, ByVal value As Double)
        Dim b = GetBytesLittle(value)
        s.Write(b, 0, 8)
    End Sub
    Public Shared Sub SetBytesLittle(ByVal s As Stream, ByVal value As Single)
        Dim b = GetBytesLittle(value)
        s.Write(b, 0, 4)
    End Sub
    Public Shared Sub SetBytesBig(ByVal s As Stream, ByVal value As Int16)
        Dim b = GetBytesBig(value)
        s.Write(b, 0, 2)
    End Sub
    Public Shared Sub SetBytesBig(ByVal s As Stream, ByVal value As Int32)
        Dim b = GetBytesBig(value)
        s.Write(b, 0, 4)
    End Sub
    Public Shared Sub SetBytesBig(ByVal s As Stream, ByVal value As Double)
        Dim b = GetBytesBig(value)
        s.Write(b, 0, 8)
    End Sub
    Public Shared Sub SetBytesBig(ByVal s As Stream, ByVal value As Single)
        Dim b = GetBytesBig(value)
        s.Write(b, 0, 4)
    End Sub

    Public Shared Sub SetBytesLittle(ByVal b() As Byte, ByVal StartIndex As Integer, ByVal value As Int16)
        Dim a = GetBytesLittle(value)
        b(StartIndex) = a(0)
        b(StartIndex + 1) = a(1)
    End Sub
    Public Shared Sub SetBytesLittle(ByVal b() As Byte, ByVal StartIndex As Integer, ByVal value As Int32)
        Dim a = GetBytesLittle(value)
        Array.Copy(a, 0, b, StartIndex, 4)
    End Sub
    Public Shared Sub SetBytesLittle(ByVal b() As Byte, ByVal StartIndex As Integer, ByVal value As Double)
        Dim a = GetBytesLittle(value)
        Array.Copy(a, 0, b, StartIndex, 8)
    End Sub
    Public Shared Sub SetBytesLittle(ByVal b() As Byte, ByVal StartIndex As Integer, ByVal value As Single)
        Dim a = GetBytesLittle(value)
        Array.Copy(a, 0, b, StartIndex, 4)
    End Sub
    Public Shared Sub SetBytesBig(ByVal b() As Byte, ByVal StartIndex As Integer, ByVal value As Int16)
        Dim a = GetBytesBig(value)
        b(StartIndex) = a(0)
        b(StartIndex + 1) = a(1)
    End Sub
    Public Shared Sub SetBytesBig(ByVal b() As Byte, ByVal StartIndex As Integer, ByVal value As Int32)
        Dim a = GetBytesBig(value)
        Array.Copy(a, 0, b, StartIndex, 4)
    End Sub
    Public Shared Sub SetBytesBig(ByVal b() As Byte, ByVal StartIndex As Integer, ByVal value As Double)
        Dim a = GetBytesBig(value)
        Array.Copy(a, 0, b, StartIndex, 8)
    End Sub
    Public Shared Sub SetBytesBig(ByVal b() As Byte, ByVal StartIndex As Integer, ByVal value As Single)
        Dim a = GetBytesBig(value)
        Array.Copy(a, 0, b, StartIndex, 4)
    End Sub

    Public Shared Function GetBytesLittle(ByVal value As Int32) As Byte()
        If BitConverter.IsLittleEndian Then
            Return BitConverter.GetBytes(value)
        Else
            Dim a As Byte
            Dim b = BitConverter.GetBytes(value)
            a = b(0) : b(0) = b(3) : b(3) = a
            a = b(1) : b(1) = b(2) : b(2) = a
            Return b
        End If
    End Function
    Public Shared Function GetBytesLittle(ByVal value As Int16) As Byte()
        If BitConverter.IsLittleEndian Then
            Return BitConverter.GetBytes(value)
        Else
            Dim b = BitConverter.GetBytes(value)
            Dim a = b(0) : b(0) = b(1) : b(1) = a
            Return b
        End If
    End Function
    Public Shared Function GetBytesLittle(ByVal value As Double) As Byte()
        If BitConverter.IsLittleEndian Then
            Return BitConverter.GetBytes(value)
        Else
            Dim a As Byte
            Dim b = BitConverter.GetBytes(value)
            a = b(0) : b(0) = b(7) : b(7) = a
            a = b(1) : b(1) = b(6) : b(6) = a
            a = b(2) : b(2) = b(5) : b(5) = a
            a = b(3) : b(3) = b(4) : b(4) = a
            Return b
        End If
    End Function
    Public Shared Function GetBytesLittle(ByVal value As Single) As Byte()
        If BitConverter.IsLittleEndian Then
            Return BitConverter.GetBytes(value)
        Else
            Dim a As Byte
            Dim b = BitConverter.GetBytes(value)
            a = b(0) : b(0) = b(3) : b(3) = a
            a = b(1) : b(1) = b(2) : b(2) = a
            Return b
        End If
    End Function
    Public Shared Function GetBytesBig(ByVal value As Int32) As Byte()
        If BitConverter.IsLittleEndian Then
            Dim a As Byte
            Dim b = BitConverter.GetBytes(value)
            a = b(0) : b(0) = b(3) : b(3) = a
            a = b(1) : b(1) = b(2) : b(2) = a
            Return b
        Else
            Return BitConverter.GetBytes(value)
        End If
    End Function
    Public Shared Function GetBytesBig(ByVal value As Int16) As Byte()
        If BitConverter.IsLittleEndian Then
            Dim b = BitConverter.GetBytes(value)
            Dim a = b(0) : b(0) = b(1) : b(1) = a
            Return b
        Else
            Return BitConverter.GetBytes(value)
        End If
    End Function
    Public Shared Function GetBytesBig(ByVal value As Double) As Byte()
        If BitConverter.IsLittleEndian Then
            Dim a As Byte
            Dim b = BitConverter.GetBytes(value)
            a = b(0) : b(0) = b(7) : b(7) = a
            a = b(1) : b(1) = b(6) : b(6) = a
            a = b(2) : b(2) = b(5) : b(5) = a
            a = b(3) : b(3) = b(4) : b(4) = a
            Return b
        Else
            Return BitConverter.GetBytes(value)
        End If
    End Function
    Public Shared Function GetBytesBig(ByVal value As Single) As Byte()
        If BitConverter.IsLittleEndian Then
            Dim a As Byte
            Dim b = BitConverter.GetBytes(value)
            a = b(0) : b(0) = b(3) : b(3) = a
            a = b(1) : b(1) = b(2) : b(2) = a
            Return b
        Else
            Return BitConverter.GetBytes(value)
        End If
    End Function
End Class

''' <summary>
''' Поле.
''' </summary>
''' <remarks>
''' Наследник класса <see cref="System.Data.DataColumn"/>.
''' </remarks>
Public Class Field
    Inherits System.Data.DataColumn

    Private _Dec As Integer
    Private _Len As Integer
    Private _Type As FieldTypes
    Friend Pos As Integer
    ''' <summary>
    ''' Общая длина, включая десятичную точку.
    ''' </summary>
    Public Property Len() As Integer
        Get
            Return _Len
        End Get
        Set(ByVal value As Integer)
            _Len = value
        End Set
    End Property
    ''' <summary>
    ''' Длина дробной части.
    ''' </summary>
    Public Property Dec() As Integer
        Get
            Return _Dec
        End Get
        Set(ByVal value As Integer)
            _Dec = value
        End Set
    End Property
    ''' <summary>
    ''' Тип поля.
    ''' </summary>
    Public ReadOnly Property FieldType() As FieldTypes
        Get
            Return _Type
        End Get
    End Property
    Friend Shared Function ConvertType(ByVal Type As FieldTypes, ByVal Dec As Integer) As System.Type
        Select Case Type
            Case FieldTypes.String
                Return GetType(System.String)
            Case FieldTypes.Float
                Return GetType(System.Double)
            Case FieldTypes.Numeric
                If Dec = 0 Then
                    Return GetType(System.Int32)
                Else
                    Return GetType(System.Double)
                End If
            Case FieldTypes.Date
                Return GetType(System.DateTime)
            Case FieldTypes.Logical
                Return GetType(System.Boolean)
            Case Else
                Return GetType(System.Byte())
        End Select
    End Function
    Public Sub New(ByVal Name As String, ByVal Type As FieldTypes, ByVal Len As Integer, Optional ByVal Dec As Integer = 0)
        MyBase.New(Name, ConvertType(Type, Dec))
        If Type = FieldTypes.Logical Then Len = 1 : Dec = 0
        If Type = FieldTypes.Date Then Len = 8 : Dec = 0
        If Type = FieldTypes.String Then Dec = 0
        If Type = FieldTypes.Memo Then Len = 10 : Dec = 0
        If Type = FieldTypes.Numeric AndAlso Len > 20 Then Len = 20
        If Type = FieldTypes.Float AndAlso Len > 20 Then Len = 20
        If Len = 0 Then
            Select Case Type
                Case FieldTypes.Numeric
                    Len = 9
                Case FieldTypes.String
                    Len = 30
                Case FieldTypes.Logical
                    Len = 1
                Case FieldTypes.Float
                    Len = 12
            End Select
        End If
        If Len < 1 Then Len = 1
        If Len > 254 Then Len = 254
        _Len = Len
        _Dec = Dec
        _Type = Type
    End Sub
    ''' <summary>
    ''' Типы полей.
    ''' </summary>
    Public Enum FieldTypes
        [String]
        [Numeric]
        [Logical]
        [Date]
        [Float]
        [Memo]
    End Enum

    Public Shadows ReadOnly Property Table() As Table
        Get
            'Return DirectCast(MyBase.Table, Table)
            Return Nothing
        End Get
    End Property
End Class

''' <summary>
''' Запись, связанная с текущей геометрией.
''' </summary>
''' <remarks>
''' Наследник класса <see cref="System.Data.DataRow"/>.
''' </remarks>
Public Class Record
    Inherits System.Data.DataRow

    Friend Sub New(ByVal rb As DataRowBuilder)
        MyBase.New(rb)
    End Sub

    Public Shadows ReadOnly Property Table() As Table
        Get
            'Return DirectCast(MyBase.Table, Table)
            Return Nothing
        End Get
    End Property
End Class

''' <summary>
''' Семантическая таблица.
''' </summary>
''' <remarks>
''' Наследник класса <see cref="System.Data.DataTable"/>.
''' </remarks>
<System.ComponentModel.DesignerCategory("Code")> _
Public Class Table
    Inherits System.Data.DataTable

    Public Shadows Function NewRow() As Record
        Return DirectCast(MyBase.NewRow, Record)
    End Function

    Protected Overrides Function NewRowFromBuilder(ByVal builder As DataRowBuilder) As DataRow
        Return New Record(builder)
    End Function
End Class

Friend Class FieldEnum
    Implements IEnumerator(Of Field)

    Public Fields As DataColumnCollection

    ' Enumerators are positioned before the first element
    ' until the first MoveNext() call.
    Dim position As Integer = -1

    Public Sub New(ByVal list As DataColumnCollection)
        Fields = list
    End Sub

    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        position = position + 1
        Return (position < Fields.Count)
    End Function

    Public Sub Reset() Implements IEnumerator.Reset
        position = -1
    End Sub

    Public ReadOnly Property Current() As Field Implements System.Collections.Generic.IEnumerator(Of Field).Current
        Get
            Try
                Return DirectCast(Fields(position), Field)
            Catch ex As IndexOutOfRangeException
                Throw New InvalidOperationException()
            End Try
        End Get
    End Property

    Public ReadOnly Property Current1() As Object Implements System.Collections.IEnumerator.Current
        Get
            Try
                Return Fields(position)
            Catch ex As IndexOutOfRangeException
                Throw New InvalidOperationException()
            End Try
        End Get
    End Property


#Region " IDisposable Support "
    Private disposedValue As Boolean = False        ' Чтобы обнаружить избыточные вызовы

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                '
            End If
        End If
        Me.disposedValue = True
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

''' <summary>
''' Коллекция полей семантической таблицы.
''' </summary>
''' <remarks>
''' Наследник класса <see cref="IEnumerable"/>.
''' </remarks>
Public Class FieldCollection : Implements IEnumerable(Of Field)

    Private _Columns As DataColumnCollection
    Friend Sub New(ByVal Columns As DataColumnCollection)
        _Columns = Columns
    End Sub
    Public Function Add(ByVal Name As String, ByVal Type As Field.FieldTypes, Optional ByVal Len As Integer = 0, Optional ByVal Dec As Integer = 0) As Field
        If Count > 127 Then
            Throw New InvalidOperationException("Field limit reached")
        End If
        Dim F = New Field(Name, Type, Len, Dec)
        _Columns.Add(F)
        Return F
    End Function
    Public Sub Add(ByVal Field As Field)
        If Count > 127 Then
            Throw New InvalidOperationException("Field limit reached")
        End If
        _Columns.Add(Field)
    End Sub
    Public Sub Remove(ByVal Name As String)
        _Columns.Remove(Name)
    End Sub
    Public Sub Remove(ByVal Field As Field)
        _Columns.Remove(Field)
    End Sub
    Public Sub RemoveAt(ByVal Index As Integer)
        _Columns.RemoveAt(Index)
    End Sub
    Public Sub Clear()
        _Columns.Clear()
    End Sub
    Public Function IndexOf(ByVal Name As String) As Integer
        Return _Columns.IndexOf(Name)
    End Function
    Public Function IndexOf(ByVal Field As Field) As Integer
        Return _Columns.IndexOf(Field)
    End Function
    Public ReadOnly Property Count() As Integer
        Get
            Return _Columns.Count
        End Get
    End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
        Return New FieldEnum(_Columns)
    End Function

    Public Function GetEnumerator1() As System.Collections.Generic.IEnumerator(Of Field) Implements System.Collections.Generic.IEnumerable(Of Field).GetEnumerator
        Return New FieldEnum(_Columns)
    End Function

    Default Public ReadOnly Property Item(ByVal Index As Integer) As Field
        Get
            Return DirectCast(_Columns(Index), Field)
        End Get
    End Property

    Default Public ReadOnly Property Item(ByVal ColumnName As String) As Field
        Get
            Return DirectCast(_Columns(ColumnName), Field)
        End Get
    End Property
End Class
