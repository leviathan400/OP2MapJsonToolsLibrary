Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Xml
Imports System.Security
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports OP2UtilityDotNet
Imports OP2UtilityDotNet.OP2Map
Imports System.Reflection

' OP2MapJsonTools / OP2MapJsonToolsLibrary
' https://github.com/leviathan400/OP2MapJsonToolsLibrary
'
' Class Library
' .NET Standard 2.0
' For shared libraries it has the broadest compatibility bridge between old and new .NET.
' I want to be able to use the library multi-platform.
'
'
' OP2MapJsonTools.vb
' Version 0.2.5.6
'
'
' Shared Library for:
' - Process Outpost 2 .map file to .json file
' - Process .json file to Outpost 2 .map file
'
'
' TO ADD:   
'           Exporting/reading Tile Groups?
'
'
' Inspired by the 'JsonMap' project.
' https://github.com/OutpostUniverse/JsonMap
' I started the JSON format as a replication of that project with flat arrays.
' But to improve it I have moved to each tile/cell row in a separate array, each on one line.
' I have padded the content for alignment for readability.
'

''' <summary>
''' The Outpost 2 map JSON format consists of the following sections:
'''
''' 1. header
''' 2. tiles
''' 3. cellTypes
''' 4. clipRect (optional)
''' 5. tileset
'''    5.1 sources
'''    5.2 tileMappings
'''    5.3 terrainTypes
'''    
''' </summary>

Public Module OP2MapJsonTools

    ''' <summary>
    ''' The Outpost 2 map JSON format consists of the following sections:
    '''
    ''' 1. header - Basic map information
    '''    - width: Number of tiles horizontally
    '''    - height: Number of tiles vertically
    '''    - map:  Filename of the original map file
    '''    - name: Name of the map
    '''    - author: Creator of the map
    '''    - notes: Additional information about the map
    '''    
    ''' 2. tiles - An array of tile mapping indices for the map's visual layer
    '''    - Linear array representing a 2D grid (stored row by row)
    '''    - Each value references an entry in the tileMappings array
    '''    - Used to determine visual appearance of each map position
    '''
    ''' 3. cellTypes - An array of cell type values for the map's gameplay layer
    '''    - Linear array representing a 2D grid (stored row by row)
    '''    - Values correspond to the CellType enum (FastPassible1, Impassible2, etc.)
    '''    - Used to determine movement costs, passability, and terrain behavior
    '''
    ''' 4. clipRect - Optional visible area definition (only if not default)
    '''    - x1, y1: Top-left coordinates
    '''    - x2, y2: Bottom-right coordinates
    '''    - Defines the visible section of the map in-game
    '''
    ''' 5. tileset - Detailed tileset information with three subsections:
    '''    
    '''    5.1 sources - Array of tileset source files
    '''        - filename: Name of the tileset file (.bmp)
    '''        - numTiles: Number of tiles in that tileset well
    '''
    '''    5.2 tileMappings - Array mapping indices to visual representation
    '''        - tilesetIndex: Which source tileset contains this tile
    '''        - tileGraphicIndex: Index within that tileset
    '''        - animationCount: Number of animation frames (0 for static)
    '''        - animationDelay: Delay between animation frames
    '''
    '''    5.3 terrainTypes - Array of terrain type definitions
    '''        - tileRange: Range of tile mapping indices using this terrain
    '''        - bulldozedTileMappingIndex: Index for bulldozed versions
    '''        - rubbleTileMappingIndex: Index for rubble versions
    '''        - tubeTileMappings: Array of indices for tube connections
    '''        - wallTileMappingIndexes: 2D array for wall configurations
    '''        - lavaTileMappingIndex: Index for lava version
    '''        - flat1/2/3: Indices for terrain variants
    '''        - tubeTileMappingIndexes: Array for tube configurations
    '''        - scorchedTileMappingIndex: Index for scorched version
    '''        - scorchedRange: Ranges for scorched variants
    '''        - unknown: Array of values with unknown purpose
    '''        
    ''' </summary>
    Public Enum JsonExportFormat
        ' Defines the format for JSON export

        Original        ' Flat arrays like the C++ implementation
        PerRow          ' With rows in separate arrays, each on one line
        PerRowPadded    ' With rows in separate arrays and numbers padded for alignment
    End Enum

    ' Where are the blank maps stored? We use these everytime we create a new .map file
    Private _blankMapsPath As String = Nothing

    ''' <summary>
    ''' Exports an Outpost 2 map file to a JSON file
    ''' </summary>
    ''' <param name="mapFileName">Path to the map file</param>
    ''' <param name="outputFileName">Path to save the JSON file</param>
    ''' <param name="format">Format to use for export (Original, PerRow, or PerRowPadded)</param>
    ''' <param name="mapName">Name of the map (for metadata)</param>
    ''' <param name="mapAuthor">Author of the map (for metadata)</param>
    Public Sub ExportMapToJsonFile(mapFileName As String, outputFileName As String, format As JsonExportFormat, mapName As String, mapAuthor As String, mapNotes As String)
        ' Read the map file
        Dim map As Map = Map.ReadMap(mapFileName)
        map.TrimTilesetSources()

        ' Get just the filename without path for the map field
        Dim mapFileNameOnly As String = Path.GetFileName(mapFileName)

        ' If mapName is empty, use the filename without extension
        If String.IsNullOrEmpty(mapName) Then
            mapName = Path.GetFileNameWithoutExtension(mapFileName)
        End If

        ' Convert to JSON based on format
        Dim json As JObject
        If format = JsonExportFormat.Original Then
            json = ConvertMapToJsonOriginal(map, mapFileNameOnly, mapName, mapAuthor, mapNotes)
        Else
            json = ConvertMapToJsonPerRow(map, mapFileNameOnly, mapName, mapAuthor, mapNotes)
        End If

        ' Write to file based on format
        If format = JsonExportFormat.Original Then
            ' Write with standard formatting
            Using writer As New StreamWriter(outputFileName, False, Encoding.UTF8)
                Dim jsonString As String = json.ToString(Newtonsoft.Json.Formatting.Indented)
                writer.Write(jsonString)
            End Using

        ElseIf format = JsonExportFormat.PerRow Then

            ' Write with custom formatting to keep rows on single lines
            Using fileStream As New FileStream(outputFileName, FileMode.Create)
                Using streamWriter As New StreamWriter(fileStream, Encoding.UTF8)
                    Using jsonWriter As New JsonTextWriter(streamWriter)
                        ' Configure the JsonTextWriter
                        jsonWriter.Formatting = Newtonsoft.Json.Formatting.Indented
                        jsonWriter.IndentChar = " "c
                        jsonWriter.Indentation = 2

                        ' Create custom serializer with special handling for arrays
                        Dim serializer As New JsonSerializer()

                        ' Write the JSON with custom formatting
                        serializer.Serialize(jsonWriter, json)
                    End Using
                End Using
            End Using

            ' Now process the file to combine array rows onto single lines
            Dim fileContent As String = File.ReadAllText(outputFileName)
            Dim processedContent As String = ProcessJsonArrays(fileContent)
            File.WriteAllText(outputFileName, processedContent)

        Else ' PerRowPadded
            ' Write with custom formatting to keep rows on single lines with padded numbers
            Using fileStream As New FileStream(outputFileName, FileMode.Create)
                Using streamWriter As New StreamWriter(fileStream, Encoding.UTF8)
                    Using jsonWriter As New JsonTextWriter(streamWriter)
                        ' Configure the JsonTextWriter
                        jsonWriter.Formatting = Newtonsoft.Json.Formatting.Indented
                        jsonWriter.IndentChar = " "c
                        jsonWriter.Indentation = 2

                        ' Create custom serializer with special handling for arrays
                        Dim serializer As New JsonSerializer()

                        ' Write the JSON with custom formatting
                        serializer.Serialize(jsonWriter, json)
                    End Using
                End Using
            End Using

            ' Now process the file to combine array rows onto single lines with padded numbers
            Dim fileContent As String = File.ReadAllText(outputFileName)
            Dim processedContent As String = ProcessJsonArraysWithPadding(fileContent)
            File.WriteAllText(outputFileName, processedContent)
        End If

        Debug.WriteLine("ExportMapToJsonFile: " & outputFileName & " - Format: " & format.ToString)
    End Sub

    ''' <summary>
    ''' Converts a Map object to a JSON object using the original flat array format
    ''' </summary>
    Private Function ConvertMapToJsonOriginal(map As Map, mapFileName As String, mapName As String, author As String, notes As String) As JObject
        Dim json As New JObject()

        '1  Add header information
        Dim width As Integer = CInt(map.WidthInTiles())
        Dim height As Integer = CInt(map.HeightInTiles())

        Dim header As New JObject()
        header.Add("width", width)
        header.Add("height", height)
        'header.Add("tileset", x)
        header.Add("map", mapFileName)
        header.Add("name", mapName)
        header.Add("author", author)
        header.Add("notes", notes)
        json.Add("header", header)

        '2  Add tile mapping data as flat array
        Dim tilesArray As New JArray()
        For i As Integer = 0 To map.TileCount() - 1
            ' Extract the x,y coordinates from the linear index
            Dim x As Integer = i Mod width
            Dim y As Integer = i \ width
            tilesArray.Add(map.GetTileMappingIndex(x, y))
        Next
        json.Add("tiles", tilesArray)

        '3  Add cell type data - using the raw enum integer values
        Dim cellTypesArray As New JArray()
        For y As Integer = 0 To height - 1
            For x As Integer = 0 To width - 1
                ' Get the cell type from the map and convert directly to integer
                Dim cellType As CellType = map.GetCellType(x, y)
                Dim cellTypeValue As Integer = CInt(cellType)
                cellTypesArray.Add(cellTypeValue)
            Next
        Next
        json.Add("cellTypes", cellTypesArray)

        '4  Add clip rect if it's not the default
        If Not IsDefaultClipRect(map.clipRect, width, height) Then
            Dim clipRectObj As New JObject()
            clipRectObj.Add("x1", map.clipRect.x1)
            clipRectObj.Add("y1", map.clipRect.y1)
            clipRectObj.Add("x2", map.clipRect.x2)
            clipRectObj.Add("y2", map.clipRect.y2)
            json.Add("clipRect", clipRectObj)
        End If

        '5  Add tileset information
        json.Add("tileset", ConvertTilesetToJson(map))

        Return json
    End Function

    ''' <summary>
    ''' Converts a Map object to a JSON object using the per-row array format
    ''' </summary>
    Private Function ConvertMapToJsonPerRow(map As Map, mapFileName As String, mapName As String, mapAuthor As String, mapNotes As String) As JObject
        Dim json As New JObject()

        '1  Add header information
        Dim width As Integer = CInt(map.WidthInTiles())
        Dim height As Integer = CInt(map.HeightInTiles())

        Dim header As New JObject()
        header.Add("width", width)
        header.Add("height", height)
        'header.Add("tileset", x)
        header.Add("map", mapFileName)
        header.Add("name", mapName)
        header.Add("author", mapAuthor)
        header.Add("notes", mapNotes)
        json.Add("header", header)

        '2  Add tile mapping data - organized by rows
        Dim tilesArray As New JArray()
        For y As Integer = 0 To height - 1
            ' Create a row array for this row of tiles
            Dim rowArray As New JArray()
            For x As Integer = 0 To width - 1
                ' Get the tile mapping index and add it to the row
                rowArray.Add(map.GetTileMappingIndex(x, y))
            Next
            ' Add the completed row to the main array
            tilesArray.Add(rowArray)
        Next
        json.Add("tiles", tilesArray)

        '3  Add cell type data - organized by rows
        Dim cellTypesArray As New JArray()
        For y As Integer = 0 To height - 1
            ' Create a row array for this row of cell types
            Dim rowArray As New JArray()
            For x As Integer = 0 To width - 1
                ' Get the cell type from the map and convert directly to integer
                Dim cellType As CellType = map.GetCellType(x, y)
                Dim cellTypeValue As Integer = CInt(cellType)

                ' Add the cell type value to the row
                rowArray.Add(cellTypeValue)
            Next
            ' Add the completed row to the main array
            cellTypesArray.Add(rowArray)
        Next
        json.Add("cellTypes", cellTypesArray)

        '4  Add clip rect if it's not the default
        If Not IsDefaultClipRect(map.clipRect, width, height) Then
            Dim clipRectObj As New JObject()
            clipRectObj.Add("x1", map.clipRect.x1)
            clipRectObj.Add("y1", map.clipRect.y1)
            clipRectObj.Add("x2", map.clipRect.x2)
            clipRectObj.Add("y2", map.clipRect.y2)
            json.Add("clipRect", clipRectObj)
        End If

        '5  Add tileset information
        json.Add("tileset", ConvertTilesetToJson(map))

        Return json
    End Function



    ''' <summary>
    ''' Processes JSON content to combine array elements onto single lines
    ''' </summary>
    Public Function ProcessJsonArrays(jsonContent As String) As String
        ' Split into lines and process
        Dim lines As String() = jsonContent.Split(New String() {Environment.NewLine}, StringSplitOptions.None)
        Dim result As New List(Of String)

        Dim inTilesSection As Boolean = False
        Dim inCellTypesSection As Boolean = False
        Dim inArrayRow As Boolean = False
        Dim currentRow As String = ""

        For i As Integer = 0 To lines.Length - 1
            Dim line As String = lines(i).TrimEnd()

            ' Check which section we're in
            If line.Contains("""tiles""") Then
                inTilesSection = True
                inCellTypesSection = False
                result.Add(line)
                Continue For
            ElseIf line.Contains("""cellTypes""") Then
                inTilesSection = False
                inCellTypesSection = True
                result.Add(line)
                Continue For
            End If

            ' Check for end of current section
            If (inTilesSection OrElse inCellTypesSection) AndAlso line = "]," Then
                inTilesSection = False
                inCellTypesSection = False
                result.Add(line)
                Continue For
            End If

            ' Handle lines within the tiles or cellTypes section
            If (inTilesSection OrElse inCellTypesSection) Then
                ' Start of an array row
                If line.TrimStart().StartsWith("[") AndAlso Not line.TrimEnd().EndsWith("]") Then
                    inArrayRow = True
                    currentRow = line.TrimEnd()
                    ' End of an array row
                ElseIf inArrayRow AndAlso (line.TrimEnd().EndsWith("],") OrElse line.TrimEnd().EndsWith("]")) Then
                    currentRow += " " + line.Trim()
                    result.Add(currentRow)
                    inArrayRow = False
                    currentRow = ""
                    ' Middle of an array row
                ElseIf inArrayRow Then
                    currentRow += " " + line.Trim()
                    ' Regular line in the section (not part of an array row)
                Else
                    result.Add(line)
                End If
                ' Lines outside the tiles or cellTypes section
            Else
                result.Add(line)
            End If
        Next

        Return String.Join(Environment.NewLine, result)
    End Function

    ''' <summary>
    ''' Processes JSON content to combine array elements onto single lines with consistent spacing
    ''' </summary>
    Public Function ProcessJsonArraysWithPadding(jsonContent As String) As String
        ' Split into lines and process
        Dim lines As String() = jsonContent.Split(New String() {Environment.NewLine}, StringSplitOptions.None)
        Dim result As New List(Of String)

        Dim inTilesSection As Boolean = False
        Dim inCellTypesSection As Boolean = False
        Dim inArrayRow As Boolean = False
        Dim currentRow As String = ""

        For i As Integer = 0 To lines.Length - 1
            Dim line As String = lines(i).TrimEnd()

            ' Check which section we're in
            If line.Contains("""tiles""") Then
                inTilesSection = True
                inCellTypesSection = False
                result.Add(line)
                Continue For
            ElseIf line.Contains("""cellTypes""") Then
                inTilesSection = False
                inCellTypesSection = True
                result.Add(line)
                Continue For
            End If

            ' Check for end of current section
            If (inTilesSection OrElse inCellTypesSection) AndAlso line = "]," Then
                inTilesSection = False
                inCellTypesSection = False
                result.Add(line)
                Continue For
            End If

            ' Handle lines within the tiles or cellTypes section
            If (inTilesSection OrElse inCellTypesSection) Then
                ' Start of an array row
                If line.TrimStart().StartsWith("[") AndAlso Not line.TrimEnd().EndsWith("]") Then
                    inArrayRow = True
                    currentRow = line.TrimEnd()
                    ' End of an array row
                ElseIf inArrayRow AndAlso (line.TrimEnd().EndsWith("],") OrElse line.TrimEnd().EndsWith("]")) Then
                    currentRow += " " + line.Trim()

                    ' Now format the row with consistent spacing based on which section we're in
                    Dim formattedRow As String
                    If inTilesSection Then
                        formattedRow = PadNumbersInRow(currentRow, 4) ' 4 spaces for tiles
                    Else ' Cell Types
                        formattedRow = PadNumbersInRow(currentRow, 2) ' 2 spaces for cell types
                    End If

                    result.Add(formattedRow)
                    inArrayRow = False
                    currentRow = ""
                    ' Middle of an array row
                ElseIf inArrayRow Then
                    currentRow += " " + line.Trim()
                    ' Regular line in the section (not part of an array row)
                Else
                    result.Add(line)
                End If
                ' Lines outside the tiles or cellTypes section
            Else
                result.Add(line)
            End If
        Next

        Return String.Join(Environment.NewLine, result)
    End Function

    ''' <summary>
    ''' Formats a row of numbers with consistent spacing based on specified padding width
    ''' </summary>
    ''' <param name="rowLine">The row line to format</param>
    ''' <param name="padWidth">Number of characters to pad to</param>
    Private Function PadNumbersInRow(rowLine As String, padWidth As Integer) As String
        ' Extract the opening bracket part
        Dim startBracketIndex As Integer = rowLine.IndexOf("[")
        Dim prefix As String = rowLine.Substring(0, startBracketIndex + 1)

        ' Extract the end bracket part
        Dim endBracketIndex As Integer = rowLine.LastIndexOf("]")
        Dim suffix As String = rowLine.Substring(endBracketIndex)

        ' Extract the content inside brackets
        Dim content As String = rowLine.Substring(startBracketIndex + 1, endBracketIndex - startBracketIndex - 1).Trim()

        ' Split the content by commas
        Dim numbers As String() = content.Split(","c)

        ' Trim and collect all numbers
        Dim trimmedNumbers As New List(Of String)
        For Each part As String In numbers
            trimmedNumbers.Add(part.Trim())
        Next

        ' Find the maximum width needed for padding
        Dim maxWidth As Integer = padWidth ' Default minimum width
        For Each numStr As String In trimmedNumbers
            If numStr.Length > maxWidth Then
                maxWidth = numStr.Length
            End If
        Next

        ' Pad each number to the maximum width
        Dim paddedNumbers As New List(Of String)
        For Each numStr As String In trimmedNumbers
            paddedNumbers.Add(numStr.PadLeft(maxWidth))
        Next

        ' Reconstruct the string with proper padding
        Return prefix & " " & String.Join(", ", paddedNumbers) & " " & suffix
    End Function


    ''' <summary>
    ''' Converts tileset information to a JSON object
    ''' </summary>
    Private Function ConvertTilesetToJson(map As Map) As JObject
        Dim json As New JObject()

        '5.1    Add sources
        Dim sourcesArray As New JArray()
        For Each source As TilesetSource In map.tilesetSources
            Dim sourceObj As New JObject()
            sourceObj.Add("filename", source.tilesetFilename)
            sourceObj.Add("numTiles", source.numTiles)
            sourcesArray.Add(sourceObj)
        Next
        json.Add("sources", sourcesArray)

        '5.2    Add tile mappings
        Dim tileMappingsArray As New JArray()
        For Each mapping As TileMapping In map.tileMappings
            Dim mappingObj As New JObject()
            mappingObj.Add("tilesetIndex", mapping.tilesetIndex)
            mappingObj.Add("tileGraphicIndex", mapping.tileGraphicIndex)
            mappingObj.Add("animationCount", mapping.animationCount)
            mappingObj.Add("animationDelay", mapping.animationDelay)
            tileMappingsArray.Add(mappingObj)
        Next
        json.Add("tileMappings", tileMappingsArray)

        '5.3    Add terrain types
        Dim terrainTypesArray As New JArray()
        For Each terrainType As TerrainType In map.terrainTypes
            Dim terrainTypeObj As New JObject()

            ' Add tile mapping range
            Dim rangeObj As New JObject()
            rangeObj.Add("start", terrainType.tileMappingRange.start)
            rangeObj.Add("end", terrainType.tileMappingRange.end)
            terrainTypeObj.Add("tileRange", rangeObj)

            terrainTypeObj.Add("bulldozedTileMappingIndex", terrainType.bulldozedTileMappingIndex)
            terrainTypeObj.Add("rubbleTileMappingIndex", terrainType.rubbleTileMappingIndex)

            ' Add tube tile mappings
            Dim tubeTileMappingsArray As New JArray()
            For Each mapping As UShort In terrainType.tubeTileMappings
                tubeTileMappingsArray.Add(mapping)
            Next
            terrainTypeObj.Add("tubeTileMappings", tubeTileMappingsArray)

            ' Add wall tile mapping indexes
            Dim wallTileMappingIndexesArray As New JArray()
            For i As Integer = 0 To terrainType.wallTileMappingIndexes.GetLength(0) - 1
                Dim wallTypeArray As New JArray()
                For j As Integer = 0 To terrainType.wallTileMappingIndexes.GetLength(1) - 1
                    wallTypeArray.Add(terrainType.wallTileMappingIndexes(i, j))
                Next
                wallTileMappingIndexesArray.Add(wallTypeArray)
            Next
            terrainTypeObj.Add("wallTileMappingIndexes", wallTileMappingIndexesArray)

            terrainTypeObj.Add("lavaTileMappingIndex", terrainType.lavaTileMappingIndex)
            terrainTypeObj.Add("flat1", terrainType.flat1)
            terrainTypeObj.Add("flat2", terrainType.flat2)
            terrainTypeObj.Add("flat3", terrainType.flat3)

            ' Add tube tile mapping indexes
            Dim tubeTileMappingIndexesArray As New JArray()
            For Each index As UShort In terrainType.tubeTileMappingIndexes
                tubeTileMappingIndexesArray.Add(index)
            Next
            terrainTypeObj.Add("tubeTileMappingIndexes", tubeTileMappingIndexesArray)

            terrainTypeObj.Add("scorchedTileMappingIndex", terrainType.scorchedTileMappingIndex)

            ' Add scorched ranges
            Dim scorchedRangeArray As New JArray()
            For Each range As Range16 In terrainType.scorchedRange
                ' Changed variable name from rangeObj to scorchedRangeObj to avoid conflict
                Dim scorchedRangeObj As New JObject()
                scorchedRangeObj.Add("start", range.start)
                scorchedRangeObj.Add("end", range.end)
                scorchedRangeArray.Add(scorchedRangeObj)
            Next
            terrainTypeObj.Add("scorchedRange", scorchedRangeArray)

            ' Add unknown data
            Dim unknownArray As New JArray()
            For Each value As Short In terrainType.unknown
                unknownArray.Add(value)
            Next
            terrainTypeObj.Add("unknown", unknownArray)

            terrainTypesArray.Add(terrainTypeObj)
        Next
        json.Add("terrainTypes", terrainTypesArray)

        Return json
    End Function




    ''' <summary>
    ''' Formats a row of numbers with consistent spacing based on specified padding width
    ''' </summary>
    ''' <param name="rowLine">The row line to format</param>
    ''' <param name="padWidth">Number of characters to pad to</param>
    Private Function FormatRowWithPadding(rowLine As String, padWidth As Integer) As String
        ' Extract the opening bracket part
        Dim startBracketIndex As Integer = rowLine.IndexOf("[")
        Dim prefix As String = rowLine.Substring(0, startBracketIndex + 1)

        ' Extract the end bracket part (may include comma)
        Dim endBracketIndex As Integer = rowLine.LastIndexOf("]")
        Dim suffix As String = rowLine.Substring(endBracketIndex)

        ' Extract the content inside brackets
        Dim content As String = rowLine.Substring(startBracketIndex + 1, endBracketIndex - startBracketIndex - 1).Trim()

        ' Split by commas to get individual numbers
        Dim numberParts As String() = content.Split(","c)

        ' Trim and collect all numbers
        Dim trimmedNumbers As New List(Of String)
        For Each part As String In numberParts
            trimmedNumbers.Add(part.Trim())
        Next

        ' Find the maximum width needed for padding
        Dim maxWidth As Integer = padWidth ' Use the provided minimum width
        For Each numStr As String In trimmedNumbers
            If numStr.Length > maxWidth Then
                maxWidth = numStr.Length
            End If
        Next

        ' Pad each number to the maximum width
        Dim paddedNumbers As New List(Of String)
        For Each numStr As String In trimmedNumbers
            paddedNumbers.Add(numStr.PadLeft(maxWidth))
        Next

        ' Reconstruct the row with padded numbers
        Return prefix & " " & String.Join(", ", paddedNumbers) & " " & suffix
    End Function

    ''' <summary>
    ''' Checks if the clip rect is the default one
    ''' </summary>
    Private Function IsDefaultClipRect(clipRect As Rect, width As Integer, height As Integer) As Boolean
        Return clipRect = GetDefaultClipRect(width, height)
    End Function

    ''' <summary>
    ''' Creates the default clip rect for a map
    ''' </summary>
    Private Function GetDefaultClipRect(width As Integer, height As Integer) As Rect
        Dim x As Integer = If(width < 512, 32, 0)
        Return New Rect(x, 0, x + width - 1, height - 2)
    End Function

#Region "Import Json to Map"


    ''' <summary>
    ''' Imports an Outpost 2 map from a JSON file and creates a .map file
    ''' </summary>
    ''' <param name="jsonFileName">Path to the JSON file</param>
    ''' <param name="outputFileName">Path to save the map file</param>
    Public Sub ImportJsonToMapFile(jsonFileName As String, outputFileName As String)

        ' MAKING NEW .MAP FILES DOES NOT MAKE A WORKING .MAP FILE.

        ' Read the JSON file
        Dim jsonText As String = File.ReadAllText(jsonFileName)
        Dim jsonObject As JObject = JObject.Parse(jsonText)

        ' Create a new Map instance
        Dim map As New Map()

        ' 1     header      Set map properties from JSON
        Dim header As JObject = DirectCast(jsonObject("header"), JObject)
        Dim width As Integer = CInt(header("width"))
        Dim height As Integer = CInt(header("height"))

        ' Setup map dimensions
        map.SetVersionTag(MapHeader.CurrentMapVersion)

        ' We need to create a properly initialized map
        ' Instead of directly setting properties, we'll initialize the tiles manually
        ' which forces the map dimensions to be correct

        ' Initialize the tiles container with the right number of tiles
        map.tiles.Clear()
        Dim tileCount As Integer = width * height
        For i As Integer = 0 To tileCount - 1
            map.tiles.Add(New Tile())
        Next


        ' Import tileset data FIRST - this is important
        Dim tileset As JObject = DirectCast(jsonObject("tileset"), JObject)
        ' 5.1   sources         Import tileset sources
        ImportTilesetSources(tileset, map)
        ' 5.2   tileMappings    Import tile mappings
        ImportTileMappings(tileset, map)
        ' 5.3   terrainTypes    Import terrain types
        ImportTerrainTypes(tileset, map)

        ' Then import tile and cell data
        ' 2     tiles       Import tile data
        ImportTiles(jsonObject, map, width, height)
        ' 3     cellTypes   Import cell types
        ImportCellTypes(jsonObject, map, width, height)

        ' Then clip rect
        ' 4     clipRect    Import clip rect if present
        If jsonObject.ContainsKey("clipRect") Then
            ImportClipRect(jsonObject, map)
        Else
            ' Set default clip rect
            Dim x As Integer = If(width < 512, 32, 0)
            map.clipRect = New Rect(x, 0, x + width - 1, height - 2)
        End If

        '' 5     tileset     Import tileset data
        'ImportTileset(jsonObject, map)


        ' Write the map file
        map.Write(outputFileName)

    End Sub

    ''' <summary>
    ''' Imports tile data from JSON
    ''' </summary>
    Private Sub ImportTiles(jsonObject As JObject, map As Map, width As Integer, height As Integer)
        ' Get the tiles array from JSON
        Dim tilesData As JToken = jsonObject("tiles")

        ' Make sure tiles container is empty
        map.tiles.Clear()

        ' Initialize the tile container to the proper size
        For i As Integer = 0 To width * height - 1
            map.tiles.Add(New Tile())
        Next

        ' Handle different array formats
        If TypeOf tilesData Is JArray Then
            Dim tilesArray As JArray = DirectCast(tilesData, JArray)

            ' Check if we have a flat array or a 2D array (PerRow format)
            If tilesArray.Count > 0 AndAlso TypeOf tilesArray(0) Is JArray Then
                ' 2D array format (PerRow or PerRowPadded)
                For y As Integer = 0 To height - 1
                    Dim rowArray As JArray = DirectCast(tilesArray(y), JArray)
                    For x As Integer = 0 To width - 1
                        map.SetTileMappingIndex(x, y, CUInt(rowArray(x)))
                    Next
                Next
            Else
                ' Flat array format (Original)
                For y As Integer = 0 To height - 1
                    For x As Integer = 0 To width - 1
                        Dim index As Integer = y * width + x
                        If index < tilesArray.Count Then
                            map.SetTileMappingIndex(x, y, CUInt(tilesArray(index)))
                        End If
                    Next
                Next
            End If
        End If
    End Sub

    ''' <summary>
    ''' Imports cell type data from JSON
    ''' </summary>
    Private Sub ImportCellTypes(jsonObject As JObject, map As Map, width As Integer, height As Integer)
        ' Get the cell types array from JSON
        Dim cellTypesData As JToken = jsonObject("cellTypes")

        ' Handle different array formats
        If TypeOf cellTypesData Is JArray Then
            Dim cellTypesArray As JArray = DirectCast(cellTypesData, JArray)

            ' Check if we have a flat array or a 2D array (PerRow format)
            If cellTypesArray.Count > 0 AndAlso TypeOf cellTypesArray(0) Is JArray Then
                ' 2D array format (PerRow or PerRowPadded)
                For y As Integer = 0 To height - 1
                    Dim rowArray As JArray = DirectCast(cellTypesArray(y), JArray)
                    For x As Integer = 0 To width - 1
                        ' Set the cell type for this location
                        map.SetCellType(CType(CInt(rowArray(x)), CellType), x, y)
                    Next
                Next
            Else
                ' Flat array format (Original)
                For y As Integer = 0 To height - 1
                    For x As Integer = 0 To width - 1
                        Dim index As Integer = y * width + x
                        map.SetCellType(CType(CInt(cellTypesArray(index)), CellType), x, y)
                    Next
                Next
            End If
        End If
    End Sub

    ''' <summary>
    ''' Imports clip rect data from JSON
    ''' </summary>
    Private Sub ImportClipRect(jsonObject As JObject, map As Map)
        Dim clipRectObj As JObject = DirectCast(jsonObject("clipRect"), JObject)

        Dim x1 As Integer = CInt(clipRectObj("x1"))
        Dim y1 As Integer = CInt(clipRectObj("y1"))
        Dim x2 As Integer = CInt(clipRectObj("x2"))
        Dim y2 As Integer = CInt(clipRectObj("y2"))

        map.clipRect = New Rect(x1, y1, x2, y2)
    End Sub


    Private Sub ImportTilesetSources(tileset As JObject, map As Map)
        Try
            ' Clear existing sources
            map.tilesetSources.Clear()

            ' Get sources array
            Dim sourcesArray As JArray = DirectCast(tileset("sources"), JArray)

            Debug.WriteLine("Found " & sourcesArray.Count & " tileset sources to import")

            ' Import each source from the JSON
            For Each sourceObj As JObject In sourcesArray
                Dim source As New TilesetSource()

                ' Set source properties directly from JSON - assuming they're already correct
                source.tilesetFilename = sourceObj("filename").ToString()
                source.numTiles = CUInt(sourceObj("numTiles"))

                ' Add to map's tilesetSources collection
                map.tilesetSources.Add(source)
                'Debug.WriteLine("Added tileset source: " & source.tilesetFilename & " with " & source.numTiles & " tiles")
            Next


            '' Pad the list to exactly 512 entries with empty sources
            'Dim currentCount = map.tilesetSources.Count
            'Debug.WriteLine("Padding tileset sources from " & currentCount & " to 512 entries")

            'For i As Integer = currentCount To 511
            '    Dim emptySource As New TilesetSource()
            '    emptySource.tilesetFilename = ""
            '    emptySource.numTiles = 0
            '    map.tilesetSources.Add(emptySource)
            'Next

            '' Verify final count
            'Debug.WriteLine("Final tileset source count: " & map.tilesetSources.Count & " (should be 512)")


            ' Show that we have 13 tilesetSources
            'Debug.WriteLine("OP2MapJsonTools.ImportTilesetSources: map.tilesetSources.Count: " & map.tilesetSources.Count)

        Catch ex As Exception
            Debug.WriteLine("ERROR in ImportTilesetSources: " & ex.Message & Environment.NewLine & ex.StackTrace)
            Throw
        End Try
    End Sub


    Private Sub ImportTerrainTypes(tileset As JObject, map As Map)
        Dim terrainTypesArray As JArray = DirectCast(tileset("terrainTypes"), JArray)

        ' Clear existing terrain types
        map.terrainTypes.Clear()

        ' Import each terrain type
        For Each terrainTypeObj As JObject In terrainTypesArray
            Dim terrainType As New TerrainType()

            ' Import tile range
            Dim rangeObj As JObject = DirectCast(terrainTypeObj("tileRange"), JObject)
            terrainType.tileMappingRange.start = CUShort(rangeObj("start"))
            terrainType.tileMappingRange.end = CUShort(rangeObj("end"))

            ' Import other properties
            terrainType.bulldozedTileMappingIndex = CUShort(terrainTypeObj("bulldozedTileMappingIndex"))
            terrainType.rubbleTileMappingIndex = CUShort(terrainTypeObj("rubbleTileMappingIndex"))

            ' Import tube tile mappings
            Dim tubeTileMappingsArray As JArray = DirectCast(terrainTypeObj("tubeTileMappings"), JArray)
            For i As Integer = 0 To Math.Min(terrainType.tubeTileMappings.Length - 1, tubeTileMappingsArray.Count - 1)
                terrainType.tubeTileMappings(i) = CUShort(tubeTileMappingsArray(i))
            Next

            ' Import wall tile mapping indexes
            Dim wallTileMappingIndexesArray As JArray = DirectCast(terrainTypeObj("wallTileMappingIndexes"), JArray)
            For i As Integer = 0 To Math.Min(terrainType.wallTileMappingIndexes.GetLength(0) - 1, wallTileMappingIndexesArray.Count - 1)
                Dim wallTypeArray As JArray = DirectCast(wallTileMappingIndexesArray(i), JArray)
                For j As Integer = 0 To Math.Min(terrainType.wallTileMappingIndexes.GetLength(1) - 1, wallTypeArray.Count - 1)
                    terrainType.wallTileMappingIndexes(i, j) = CUShort(wallTypeArray(j))
                Next
            Next

            ' Import other indices
            terrainType.lavaTileMappingIndex = CUShort(terrainTypeObj("lavaTileMappingIndex"))
            terrainType.flat1 = CUShort(terrainTypeObj("flat1"))
            terrainType.flat2 = CUShort(terrainTypeObj("flat2"))
            terrainType.flat3 = CUShort(terrainTypeObj("flat3"))

            ' Import tube tile mapping indexes
            Dim tubeTileMappingIndexesArray As JArray = DirectCast(terrainTypeObj("tubeTileMappingIndexes"), JArray)
            For i As Integer = 0 To Math.Min(terrainType.tubeTileMappingIndexes.Length - 1, tubeTileMappingIndexesArray.Count - 1)
                terrainType.tubeTileMappingIndexes(i) = CUShort(tubeTileMappingIndexesArray(i))
            Next

            terrainType.scorchedTileMappingIndex = CUShort(terrainTypeObj("scorchedTileMappingIndex"))

            ' Import scorched ranges
            Dim scorchedRangeArray As JArray = DirectCast(terrainTypeObj("scorchedRange"), JArray)
            For i As Integer = 0 To Math.Min(terrainType.scorchedRange.Length - 1, scorchedRangeArray.Count - 1)
                Dim scorchedRangeObj As JObject = DirectCast(scorchedRangeArray(i), JObject)
                terrainType.scorchedRange(i).start = CUShort(scorchedRangeObj("start"))
                terrainType.scorchedRange(i).end = CUShort(scorchedRangeObj("end"))
            Next

            ' Import unknown data
            Dim unknownArray As JArray = DirectCast(terrainTypeObj("unknown"), JArray)
            For i As Integer = 0 To Math.Min(terrainType.unknown.Length - 1, unknownArray.Count - 1)
                terrainType.unknown(i) = CShort(unknownArray(i))
            Next

            map.terrainTypes.Add(terrainType)
        Next
    End Sub

    ''' <summary>
    ''' Imports tileset mapping data from JSON
    ''' </summary>
    Private Sub ImportTileMappings(tileset As JObject, map As Map)
        Dim tileMappingsArray As JArray = DirectCast(tileset("tileMappings"), JArray)

        ' Clear existing tile mappings
        map.tileMappings.Clear()

        ' Import each tile mapping
        For Each mappingObj As JObject In tileMappingsArray
            Dim mapping As New TileMapping()
            mapping.tilesetIndex = CUShort(mappingObj("tilesetIndex"))
            mapping.tileGraphicIndex = CUShort(mappingObj("tileGraphicIndex"))
            mapping.animationCount = CUShort(mappingObj("animationCount"))
            mapping.animationDelay = CUShort(mappingObj("animationDelay"))

            map.tileMappings.Add(mapping)
        Next
    End Sub




#End Region


    'Public Sub CreateStandardMapFromScratch(outputFileName As String)

    '    ' DOES NOT CREATE A VALID .MAP FILE

    '    Try
    '        ' Create a new map
    '        Dim map As New Map()

    '        ' Set version tag
    '        map.SetVersionTag(MapHeader.CurrentMapVersion)

    '        ' Set dimensions (small test map)
    '        Dim width As Integer = 64
    '        Dim height As Integer = 64

    '        ' Initialize tiles collection
    '        map.tiles.Clear()
    '        For i As Integer = 0 To width * height - 1
    '            map.tiles.Add(New Tile())
    '        Next

    '        ' Set default clip rect
    '        Dim clipX As Integer = If(width < 512, 32, 0)
    '        map.clipRect = New Rect(clipX, 0, clipX + width - 1, height - 2)

    '        ' Clear and initialize tileset sources with standard OP2 wells
    '        map.tilesetSources.Clear()

    '        ' Add the standard 13 wells with correct tile counts
    '        AddTilesetSource(map, "well0000", 1)
    '        AddTilesetSource(map, "well0001", 269)
    '        AddTilesetSource(map, "well0002", 163)
    '        AddTilesetSource(map, "well0003", 6)
    '        AddTilesetSource(map, "well0004", 359)
    '        AddTilesetSource(map, "well0005", 147)
    '        AddTilesetSource(map, "well0006", 54)
    '        AddTilesetSource(map, "well0007", 207)
    '        AddTilesetSource(map, "well0008", 347)
    '        AddTilesetSource(map, "well0009", 141)
    '        AddTilesetSource(map, "well0010", 96)
    '        AddTilesetSource(map, "well0011", 150)
    '        AddTilesetSource(map, "well0012", 72)

    '        ' Pad to 512 entries
    '        Dim currentCount = map.tilesetSources.Count
    '        For i As Integer = currentCount To 511
    '            Dim emptySource As New TilesetSource()
    '            emptySource.tilesetFilename = ""
    '            emptySource.numTiles = 0
    '            map.tilesetSources.Add(emptySource)
    '        Next

    '        Debug.WriteLine("Added " & currentCount & " standard wells and padded to " & map.tilesetSources.Count & " entries")

    '        ' Create a minimal terrain type
    '        Dim terrainType As New TerrainType()
    '        terrainType.tileMappingRange.start = 0
    '        terrainType.tileMappingRange.end = 10  ' Give it some range
    '        map.terrainTypes.Add(terrainType)

    '        ' Create some basic tile mappings
    '        For i As Integer = 0 To 10
    '            Dim tileMapping As New TileMapping()
    '            tileMapping.tilesetIndex = 0  ' Use the first tileset
    '            tileMapping.tileGraphicIndex = 0
    '            tileMapping.animationCount = 0
    '            tileMapping.animationDelay = 0
    '            map.tileMappings.Add(tileMapping)
    '        Next

    '        ' Set all tiles to use the first tile mapping
    '        For y As Integer = 0 To height - 1
    '            For x As Integer = 0 To width - 1
    '                map.SetTileMappingIndex(x, y, 0)
    '                map.SetCellType(CellType.FastPassible1, x, y)
    '            Next
    '        Next

    '        ' Trim tileset sources
    '        map.TrimTilesetSources()
    '        Debug.WriteLine("Tileset sources count after trimming: " & map.tilesetSources.Count)

    '        ' Make sure it's still properly padded to 512
    '        currentCount = map.tilesetSources.Count
    '        If currentCount < 512 Then
    '            Debug.WriteLine("Re-padding tileset sources to 512 after trimming")
    '            For i As Integer = currentCount To 511
    '                Dim emptySource As New TilesetSource()
    '                emptySource.tilesetFilename = ""
    '                emptySource.numTiles = 0
    '                map.tilesetSources.Add(emptySource)
    '            Next
    '        End If

    '        ' Write the map
    '        Debug.WriteLine("Writing map to file: " & outputFileName)
    '        map.Write(outputFileName)
    '        Debug.WriteLine("Standard map created successfully at: " & outputFileName)

    '    Catch ex As Exception
    '        Debug.WriteLine("ERROR in CreateStandardMapFromScratch: " & ex.Message & Environment.NewLine & ex.StackTrace)
    '        Throw
    '    End Try
    'End Sub


    ' Helper function to add a tileset source
    Private Sub AddTilesetSource(map As Map, filename As String, numTiles As UInteger)
        Dim source As New TilesetSource()
        source.tilesetFilename = filename
        source.numTiles = numTiles
        map.tilesetSources.Add(source)
        Debug.WriteLine("Added tileset source: " & filename & " with " & numTiles & " tiles")
    End Sub




    Public Sub UpdateAndSaveMap(map As Map, outputFilePath As String)
        Try
            Debug.WriteLine("Updating map tiles and cell types...")

            ' Get map dimensions
            Dim width As Integer = CInt(map.WidthInTiles())
            Dim height As Integer = CInt(map.HeightInTiles())
            Debug.WriteLine("Map dimensions: " & width & "x" & height)

            ' Update all tiles to use tile mapping index 0 and cellType FastPassible1
            For y As Integer = 0 To height - 1
                For x As Integer = 0 To width - 1

                    ' Set tile mapping index to 0
                    map.SetTileMappingIndex(x, y, 0)

                    ' Set cell type to a basic passable type
                    map.SetCellType(CellType.FastPassible1, x, y)

                Next
            Next
            'Debug.WriteLine("All tiles updated to mapping index 0")
            'Debug.WriteLine("All cellType tiles updated to FastPassible1")

            'Debug.WriteLine("All tiles updated to mapping index 0 and FastPassible1")

            '' Update tileset sources from "well00" to "grnwld"
            'Debug.WriteLine("Updating tileset sources from 'well00' to 'grnwld'...")
            'Dim updatedCount As Integer = 0
            'For i As Integer = 0 To map.tilesetSources.Count - 1
            '    Dim source As TilesetSource = map.tilesetSources(i)

            '    ' Only process non-empty sources
            '    If Not String.IsNullOrEmpty(source.tilesetFilename) Then
            '        ' Simply replace "well00" with "grnwld" in the filename
            '        If source.tilesetFilename.Contains("well00") Then
            '            source.tilesetFilename = source.tilesetFilename.Replace("well00", "grnwld")
            '            updatedCount += 1
            '        End If
            '    End If
            'Next
            'Debug.WriteLine($"Updated {updatedCount} tileset sources from 'well00' to 'grnwld'")


            ' Remove all tile groups
            Debug.WriteLine("Removing " & map.tileGroups.Count & " tile groups...")
            map.tileGroups.Clear()
            Debug.WriteLine("All tile groups removed")


            '' Add a tile group named "tester"
            'Debug.WriteLine("Adding tile group 'tester'...")

            '' Create a new tile group
            'Dim tileGroup As New TileGroup()
            'tileGroup.name = "tester"
            'tileGroup.tileWidth = 1
            'tileGroup.tileHeight = 1

            '' Add the tile mapping index 0 to the group
            'tileGroup.mappingIndices.Add(0)

            '' Add the tile group to the map
            'map.tileGroups.Add(tileGroup)

            'Debug.WriteLine("Added tile group 'tester' with 1 tile mapping index (0)")




            ' Remove empty tileset sources
            'Debug.WriteLine("Removing empty tileset sources...")
            'Debug.WriteLine("Total tileset sources before cleaning: " & map.tilesetSources.Count)

            ' First, identify which indices contain non-empty tileset sources
            Dim nonEmptyIndices As New List(Of Integer)()
            For i As Integer = 0 To map.tilesetSources.Count - 1
                Dim source As TilesetSource = map.tilesetSources(i)
                If Not String.IsNullOrEmpty(source.tilesetFilename) AndAlso source.numTiles > 0 Then
                    nonEmptyIndices.Add(i)
                End If
            Next

            'Debug.WriteLine("Found " & nonEmptyIndices.Count & " non-empty tileset sources")

            ' Create a new list with only the non-empty sources
            Dim cleanedSources As New List(Of TilesetSource)()
            For Each index As Integer In nonEmptyIndices
                cleanedSources.Add(map.tilesetSources(index))
            Next

            ' Replace the original collection with our cleaned version
            map.tilesetSources.Clear()
            For Each source As TilesetSource In cleanedSources
                map.tilesetSources.Add(source)
            Next

            Debug.WriteLine("Removed " & (512 - cleanedSources.Count) & " empty tileset sources")
            'Debug.WriteLine("Total tileset sources after cleaning: " & map.tilesetSources.Count)




            ' Save the updated map to the specified path
            Debug.WriteLine("Saving map to: " & outputFilePath)
            map.Write(outputFilePath)

            Debug.WriteLine("Map saved successfully")

        Catch ex As Exception
            Debug.WriteLine("ERROR in UpdateAndSaveMap: " & ex.Message & Environment.NewLine & ex.StackTrace)
            Throw
        End Try
    End Sub

    Public Sub UpdateMap_CreateBlankMap(map As Map, outputFilePath As String)
        ' - Create a blank map with the dimensions of the loaded map
        '
        ' 1. Update all tiles to mapping index 0
        ' 2. Update all cellType tiles to FastPassible1
        ' 3. Remove empty tileset sources (Reduce from 512 to 13)
        ' 4. Remove all tile groups
        '
        ' 5. Save map to outputFilePath

        Try
            Debug.WriteLine("Creating blank map from loaded map")
            'Debug.WriteLine("Updating map tiles and cell types...")

            ' Get map dimensions
            Dim width As Integer = CInt(map.WidthInTiles())
            Dim height As Integer = CInt(map.HeightInTiles())
            Debug.WriteLine("Map dimensions: " & width & "x" & height)

            ' Update all tiles to use tile mapping index 0 and cellType FastPassible1
            For y As Integer = 0 To height - 1
                For x As Integer = 0 To width - 1

                    ' 1     Set tile mapping index to 0
                    map.SetTileMappingIndex(x, y, 0)

                    ' 2     Set cell type to a basic passable type
                    map.SetCellType(CellType.FastPassible1, x, y)

                Next
            Next
            'Debug.WriteLine("All tiles updated to mapping index 0")
            'Debug.WriteLine("All cellType tiles updated to FastPassible1")

            'Debug.WriteLine("All tiles updated to mapping index 0 and FastPassible1")

            '' Update tileset sources from "well00" to "grnwld"
            'Debug.WriteLine("Updating tileset sources from 'well00' to 'grnwld'...")
            'Dim updatedCount As Integer = 0
            'For i As Integer = 0 To map.tilesetSources.Count - 1
            '    Dim source As TilesetSource = map.tilesetSources(i)

            '    ' Only process non-empty sources
            '    If Not String.IsNullOrEmpty(source.tilesetFilename) Then
            '        ' Simply replace "well00" with "grnwld" in the filename
            '        If source.tilesetFilename.Contains("well00") Then
            '            source.tilesetFilename = source.tilesetFilename.Replace("well00", "grnwld")
            '            updatedCount += 1
            '        End If
            '    End If
            'Next
            'Debug.WriteLine($"Updated {updatedCount} tileset sources from 'well00' to 'grnwld'")





            '' Add a tile group named "tester"
            'Debug.WriteLine("Adding tile group 'tester'...")

            '' Create a new tile group
            'Dim tileGroup As New TileGroup()
            'tileGroup.name = "tester"
            'tileGroup.tileWidth = 1
            'tileGroup.tileHeight = 1

            '' Add the tile mapping index 0 to the group
            'tileGroup.mappingIndices.Add(0)

            '' Add the tile group to the map
            'map.tileGroups.Add(tileGroup)

            'Debug.WriteLine("Added tile group 'tester' with 1 tile mapping index (0)")




            ' 3     Remove empty tileset sources
            'Debug.WriteLine("Removing empty tileset sources...")
            'Debug.WriteLine("Total tileset sources before cleaning: " & map.tilesetSources.Count)
            ' First, identify which indices contain non-empty tileset sources
            Dim nonEmptyIndices As New List(Of Integer)()
            For i As Integer = 0 To map.tilesetSources.Count - 1
                Dim source As TilesetSource = map.tilesetSources(i)
                If Not String.IsNullOrEmpty(source.tilesetFilename) AndAlso source.numTiles > 0 Then
                    nonEmptyIndices.Add(i)
                End If
            Next
            'Debug.WriteLine("Found " & nonEmptyIndices.Count & " non-empty tileset sources")
            ' Create a new list with only the non-empty sources
            Dim cleanedSources As New List(Of TilesetSource)()
            For Each index As Integer In nonEmptyIndices
                cleanedSources.Add(map.tilesetSources(index))
            Next
            ' Replace the original collection with our cleaned version
            map.tilesetSources.Clear()
            For Each source As TilesetSource In cleanedSources
                map.tilesetSources.Add(source)
            Next
            Debug.WriteLine("Removed " & (512 - cleanedSources.Count) & " empty tileset sources")
            Debug.WriteLine("Total tileset sources after cleaning: " & map.tilesetSources.Count)


            ' 4     Remove all tile groups
            'Debug.WriteLine("Removing " & map.tileGroups.Count & " tile groups...")
            map.tileGroups.Clear()
            'Debug.WriteLine("All tile groups removed")

            ' 5     Save the updated map to the specified path
            Debug.WriteLine("Saving map to: " & outputFilePath)
            map.Write(outputFilePath)

            Debug.WriteLine("Map saved successfully")

        Catch ex As Exception
            Debug.WriteLine("ERROR in UpdateAndSaveMap: " & ex.Message & Environment.NewLine & ex.StackTrace)
            Throw
        End Try
    End Sub



#Region "JSON to MAP"

    Public Sub ExportMapFile(jsonFilePath As String, outputMapFilePath As String)
        ' jsonFilePath - The full file path to the json file containing the map data
        ' outputMapFilePath - The full file path of the map file that we want to save

        ' Create .map file from json data
        Dim map As Map
        Debug.WriteLine("OP2MapJsonTools.ExportMapFile: Creating .map from .json data")

        ' Read and parse the JSON file
        Dim jsonText As String = File.ReadAllText(jsonFilePath)
        Dim jsonObject As JObject = JObject.Parse(jsonText)
        Dim mapWidth As Int16
        Dim mapHeight As Int16

        'Dim mapTileset As String ' Unsure whether to include tileset in the header as we can find the same data from the tileset sources section
        Dim mapFileName As String
        Dim mapName As String
        Dim mapAuthor As String
        Dim mapNotes As String
        Try
            ' Extract header information - Get map dimensions from the json file
            Dim header As JObject = DirectCast(jsonObject("header"), JObject)
            mapWidth = CInt(header("width"))            ' Get map width
            mapHeight = CInt(header("height"))          ' Get map height
            'mapTileset = header("tileset").ToString()   ' Get map tileset name
            mapFileName = header("map").ToString()      ' Get map file name
            mapName = header("name").ToString()         ' Get map name
            mapAuthor = header("author").ToString()     ' Get author name
            mapNotes = header("notes").ToString()       ' Get map notes
        Catch ex As Exception

        End Try

        Try
            'AppendToConsole("Starting map export with blank map template...")

            ' Determine which blank map to use based on dimensions from JSON
            Dim blankMapFileName As String = "blank_" & mapWidth & "x" & mapHeight & ".map"
            ' Blank Maps (10)
            ' blank_64x64.map
            ' blank_64x128.map
            ' blank_64x256.map
            ' blank_128x64.map
            ' blank_128x128.map
            ' blank_128x256.map
            ' blank_256x64.map
            ' blank_256x128.map
            ' blank_256x256.map
            ' blank_512x256.map
            'Dim blankMapFileName As String = GetBlankMapFileName(mapWidth, mapHeight)      ' Not really needed?

            ' Check if we found a matching template
            If String.IsNullOrEmpty(blankMapFileName) Then
                Dim errorMsg As String = $"ERROR: No blank map template found for dimensions {mapWidth}x{mapHeight}"
                'AppendToConsole(errorMsg)
                Debug.WriteLine(errorMsg)
                'MessageBox.Show($"Could not find blank map template: blank_{mapWidth}x{mapHeight}.map", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            Dim blankMapsPath As String = EnsureBlankMapsExtracted()
            Dim blankMapPath As String = Path.Combine(blankMapsPath, blankMapFileName)
            'AppendToConsole("Using blank map template: " & blankMapFileName)

            ' Load the blank map
            map = Map.ReadMap(blankMapPath)
            'AppendToConsole("Blank map loaded successfully: " & currentMap.WidthInTiles & "x" & currentMap.HeightInTiles)


            ' Process tiles 
            Debug.WriteLine("Importing tiles data...")
            Dim tilesData As JToken = jsonObject("tiles")

            If TypeOf tilesData Is JArray Then
                Dim tilesArray As JArray = DirectCast(tilesData, JArray)

                ' Per-row format
                Debug.WriteLine("Processing per-row format for tiles")
                For y As Integer = 0 To mapHeight - 1
                    If y < tilesArray.Count Then
                        Dim rowArray As JArray = DirectCast(tilesArray(y), JArray)
                        For x As Integer = 0 To mapWidth - 1
                            If x < rowArray.Count Then
                                map.SetTileMappingIndex(x, y, CUInt(rowArray(x)))
                            End If
                        Next
                    End If
                Next
            End If

            ' Process cell types 
            Debug.WriteLine("Importing cell types data...")
            Dim cellTypesData As JToken = jsonObject("cellTypes")

            If TypeOf cellTypesData Is JArray Then
                Dim cellTypesArray As JArray = DirectCast(cellTypesData, JArray)

                ' Per-row format
                Debug.WriteLine("Processing per-row format for cell types")
                For y As Integer = 0 To mapHeight - 1
                    If y < cellTypesArray.Count Then
                        Dim rowArray As JArray = DirectCast(cellTypesArray(y), JArray)
                        For x As Integer = 0 To mapWidth - 1
                            If x < rowArray.Count Then
                                map.SetCellType(CType(CInt(rowArray(x)), CellType), x, y)
                            End If
                        Next
                    End If
                Next
            End If

            ' Set clip rect if specified in JSON
            If jsonObject.ContainsKey("clipRect") Then
                Debug.WriteLine("Using custom clip rect from JSON...")
                Dim clipRectObj As JObject = DirectCast(jsonObject("clipRect"), JObject)
                map.clipRect = New Rect(
                CInt(clipRectObj("x1")),
                CInt(clipRectObj("y1")),
                CInt(clipRectObj("x2")),
                CInt(clipRectObj("y2"))
            )
            Else
                ' Set default clip rect
                Debug.WriteLine("Using default clip rect...")
                Dim x As Integer = If(mapWidth < 512, 32, 0)
                map.clipRect = New Rect(x, 0, x + mapWidth - 1, mapHeight - 2)
            End If

            ' Process only the tileset sources if available
            If jsonObject.ContainsKey("tileset") Then
                ImportTilesetData(jsonObject, map)
            End If

            ' Clear any existing tile groups to ensure they're empty
            map.tileGroups.Clear()

            ' Save the updated map to disk
            Debug.WriteLine("Saving map to: " & outputMapFilePath)
            map.Write(outputMapFilePath)
            Debug.WriteLine("Map export complete.")

            TidyUp()

        Catch ex As Exception
            Debug.WriteLine("ERROR in ExportMapFile: " & ex.Message & Environment.NewLine & ex.StackTrace)
            Throw
        End Try
    End Sub


    Private Sub ImportTilesetData(jsonObject As JObject, map As Map)
        Try
            Debug.WriteLine("Importing tileset sources only...")
            Dim tileset As JObject = DirectCast(jsonObject("tileset"), JObject)

            ' Import only the sources
            If tileset.ContainsKey("sources") Then
                Debug.WriteLine("Importing tileset sources...")
                ImportTilesetSources(tileset, map)
            End If

            ' Note: We're intentionally NOT importing tileMappings and terrainTypes
            ' as we want to keep the ones from the blank map template
            Debug.WriteLine("Skipping import of tileMappings and terrainTypes (using template values)")

        Catch ex As Exception
            Debug.WriteLine("ERROR in ImportTilesetData: " & ex.Message & Environment.NewLine & ex.StackTrace)
            Throw
        End Try
    End Sub


#End Region




#Region "Blank Maps Management"

    '''' <summary>
    '''' Gets the file name for the blank map
    '''' </summary>
    'Private Function GetBlankMapFileName(width As Integer, height As Integer) As String
    '    ' Only return an exact match for the dimensions
    '    Dim exactFileName As String = $"blank_{width}x{height}.map"

    '    ' Blank Maps (10)
    '    ' blank_64x64.map
    '    ' blank_64x128.map
    '    ' blank_64x256.map
    '    ' blank_128x64.map
    '    ' blank_128x128.map
    '    ' blank_128x256.map
    '    ' blank_256x64.map
    '    ' blank_256x128.map
    '    ' blank_256x256.map
    '    ' blank_512x256.map

    '    If File.Exists(Path.Combine(BlankMapsPath, exactFileName)) Then
    '        Return exactFileName
    '    Else
    '        ' No exact match found - return empty string to indicate error
    '        Return String.Empty
    '    End If
    'End Function

    ''' <summary>
    ''' Ensures blank maps are extracted and returns the path to the folder
    ''' </summary>
    ''' <returns>Path to the blank maps folder</returns>
    Private Function EnsureBlankMapsExtracted() As String
        If _blankMapsPath IsNot Nothing Then
            Return _blankMapsPath
        End If

        Try
            ' Get the library startup directory
            Dim libraryPath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            If String.IsNullOrEmpty(libraryPath) Then
                libraryPath = Directory.GetCurrentDirectory()
            End If

            ' Determine folder name based on platform
            Dim folderName As String
            If RuntimeInformation.IsOSPlatform(OSPlatform.Windows) Then
                folderName = "blankmaps"
            Else
                folderName = ".blankmaps" ' Hidden on Linux/Mac
            End If

            Dim blankMapsFolder As String = Path.Combine(libraryPath, folderName)

            ' Check if folder exists and has files
            If Not Directory.Exists(blankMapsFolder) OrElse Directory.GetFiles(blankMapsFolder, "*.map").Length = 0 Then
                ' Extract the blank maps
                Debug.WriteLine("Blank maps folder not found or empty - extracting from zip...")
                ExtractBlankMapsFromZip(blankMapsFolder)

                ' Set as hidden on Windows
                If RuntimeInformation.IsOSPlatform(OSPlatform.Windows) Then
                    Try
                        Dim dirInfo As New DirectoryInfo(blankMapsFolder)
                        dirInfo.Attributes = dirInfo.Attributes Or FileAttributes.Hidden
                        Debug.WriteLine("Blank maps folder set as hidden")
                    Catch ex As Exception
                        Debug.WriteLine("Could not set folder as hidden: " & ex.Message)
                    End Try
                End If
            Else
                ' Folder exists with files - no extraction needed
                Dim mapCount As Integer = Directory.GetFiles(blankMapsFolder, "*.map").Length
                Debug.WriteLine($"Blank maps folder found with {mapCount} files - skipping extraction")
            End If

            _blankMapsPath = blankMapsFolder
            Debug.WriteLine("Blank maps folder: " & _blankMapsPath)
            Return _blankMapsPath

        Catch ex As Exception
            Debug.WriteLine("ERROR ensuring blank maps extracted: " & ex.Message)
            Throw New Exception("Failed to extract blank maps", ex)
        End Try
    End Function

    ''' <summary>
    ''' Extracts blank maps from embedded zip to the specified folder
    ''' </summary>
    ''' <param name="targetFolder">Folder to extract blank maps to</param>
    Private Sub ExtractBlankMapsFromZip(targetFolder As String)
        Try
            ' Find the embedded zip resource
            Dim assembly As Assembly = Assembly.GetExecutingAssembly()
            Dim resourceNames As String() = assembly.GetManifestResourceNames()

            ' Debug: List all embedded resources
            'Debug.WriteLine("All embedded resources:")
            'For Each name As String In resourceNames
            '    Debug.WriteLine("  - " & name)
            'Next

            Dim zipResourceName As String = Nothing
            For Each resourceName As String In resourceNames
                If resourceName.ToLower().EndsWith(".zip") AndAlso
                   (resourceName.ToLower().Contains("blank") OrElse resourceName.ToLower().Contains("maps")) Then
                    zipResourceName = resourceName
                    Exit For
                End If
            Next

            If String.IsNullOrEmpty(zipResourceName) Then
                Throw New Exception("Could not find embedded blankmaps.zip resource")
            End If

            'Debug.WriteLine("Using zip resource: " & zipResourceName)          'OP2MapJsonTools.blankmaps.zip

            ' Create target directory
            If Not Directory.Exists(targetFolder) Then
                Directory.CreateDirectory(targetFolder)
            End If

            ' Extract the zip - use the FULL resource name
            Using resourceStream As Stream = assembly.GetManifestResourceStream(zipResourceName)
                If resourceStream Is Nothing Then
                    Throw New Exception($"Could not load resource stream for: {zipResourceName}")
                End If
                'Debug.WriteLine("Successfully opened resource stream")

                Using zipArchive As New ZipArchive(resourceStream, ZipArchiveMode.Read)
                    Dim extractedCount As Integer = 0
                    'Debug.WriteLine($"Zip archive contains {zipArchive.Entries.Count} entries")

                    For Each entry As ZipArchiveEntry In zipArchive.Entries
                        'Debug.WriteLine($"Processing zip entry: {entry.Name}")
                        If entry.Name.EndsWith(".map") AndAlso entry.Name.StartsWith("blank_") Then
                            Dim targetPath As String = Path.Combine(targetFolder, entry.Name)

                            Using entryStream As Stream = entry.Open()
                                Using fileStream As New FileStream(targetPath, FileMode.Create, FileAccess.Write)
                                    entryStream.CopyTo(fileStream)
                                End Using
                            End Using

                            extractedCount += 1
                            'Debug.WriteLine("Extracted: " & entry.Name)
                        End If
                    Next

                    Debug.WriteLine($"Successfully extracted {extractedCount} blank map files")

                    If extractedCount = 0 Then
                        Throw New Exception("No blank map files were found in the zip archive")
                    End If
                End Using
            End Using

        Catch ex As Exception
            Debug.WriteLine("ERROR extracting blank maps from zip: " & ex.Message)
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' Removes the extracted blank maps folder and all its contents
    ''' </summary>
    Private Sub TidyUp()
        Try
            'Debug.WriteLine("TidyUp: Starting cleanup of blank maps...")

            ' Reset the cached path
            Dim folderToDelete As String = _blankMapsPath
            _blankMapsPath = Nothing

            ' If we don't have a cached path, try to determine it
            If String.IsNullOrEmpty(folderToDelete) Then
                Dim libraryPath As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                If String.IsNullOrEmpty(libraryPath) Then
                    libraryPath = Directory.GetCurrentDirectory()
                End If

                ' Check both possible folder names
                Dim windowsFolder As String = Path.Combine(libraryPath, "blankmaps")
                Dim unixFolder As String = Path.Combine(libraryPath, ".blankmaps")

                If Directory.Exists(windowsFolder) Then
                    folderToDelete = windowsFolder
                ElseIf Directory.Exists(unixFolder) Then
                    folderToDelete = unixFolder
                End If
            End If

            ' Delete the folder if it exists
            If Not String.IsNullOrEmpty(folderToDelete) AndAlso Directory.Exists(folderToDelete) Then
                ' Remove hidden attribute if present (Windows)
                If RuntimeInformation.IsOSPlatform(OSPlatform.Windows) Then
                    Try
                        Dim dirInfo As New DirectoryInfo(folderToDelete)
                        If (dirInfo.Attributes And FileAttributes.Hidden) = FileAttributes.Hidden Then
                            dirInfo.Attributes = dirInfo.Attributes And Not FileAttributes.Hidden
                            'Debug.WriteLine("TidyUp: Removed hidden attribute from folder")
                        End If
                    Catch ex As Exception
                        Debug.WriteLine("TidyUp: Could not remove hidden attribute: " & ex.Message)
                    End Try
                End If

                ' Count files before deletion
                Dim fileCount As Integer = Directory.GetFiles(folderToDelete, "*.map").Length

                ' Delete the entire folder and its contents
                Directory.Delete(folderToDelete, True)

                Debug.WriteLine($"TidyUp: Successfully deleted folder '{Path.GetFileName(folderToDelete)}' with {fileCount} files")
            Else
                Debug.WriteLine("TidyUp: No blank maps folder found to delete")
            End If

        Catch ex As Exception
            Debug.WriteLine("TidyUp ERROR: " & ex.Message)
            Throw New Exception("Failed to tidy up blank maps folder", ex)
        End Try
    End Sub

#End Region



    ' Idea to use tile group area for metadata (strings)
    '#Region "Metadata in Tile Groups"

    '    ' Add metadata as tile groups
    '    Public Sub AddMetadataAsTileGroups(map As Map, author As String, description As String, version As String)
    '        Debug.WriteLine("Adding metadata as tile groups...")

    '        ' Clear existing tile groups if needed
    '        map.tileGroups.Clear()

    '        ' Create author tile group
    '        If Not String.IsNullOrEmpty(author) Then
    '            Dim authorGroup As New TileGroup()
    '            authorGroup.name = "META_AUTHOR_" & author
    '            authorGroup.tileWidth = 1
    '            authorGroup.tileHeight = 1
    '            authorGroup.mappingIndices.Add(0)
    '            map.tileGroups.Add(authorGroup)
    '            Debug.WriteLine("Added metadata for author: " & author)
    '        End If

    '        ' Create version tile group
    '        If Not String.IsNullOrEmpty(version) Then
    '            Dim versionGroup As New TileGroup()
    '            versionGroup.name = "META_VERSION_" & version
    '            versionGroup.tileWidth = 1
    '            versionGroup.tileHeight = 1
    '            versionGroup.mappingIndices.Add(0)
    '            map.tileGroups.Add(versionGroup)
    '            Debug.WriteLine("Added metadata for version: " & version)
    '        End If

    '        ' Create description tile group - may need to be truncated if too long
    '        If Not String.IsNullOrEmpty(description) Then
    '            ' Limit description length if needed
    '            Dim maxLength As Integer = 200
    '            Dim safeDescription As String = If(description.Length > maxLength,
    '                                           description.Substring(0, maxLength) & "...",
    '                                           description)

    '            Dim descGroup As New TileGroup()
    '            descGroup.name = "META_DESC_" & safeDescription
    '            descGroup.tileWidth = 1
    '            descGroup.tileHeight = 1
    '            descGroup.mappingIndices.Add(0)
    '            map.tileGroups.Add(descGroup)
    '            Debug.WriteLine("Added metadata for description")
    '        End If

    '        Debug.WriteLine("Metadata tile groups added successfully")
    '    End Sub


    '    ' Extract metadata from tile groups
    '    Public Function ExtractMetadataFromTileGroups(map As Map) As Dictionary(Of String, String)
    '        Dim metadata As New Dictionary(Of String, String)

    '        For Each group As TileGroup In map.tileGroups
    '            If group.name.StartsWith("META_AUTHOR_") Then
    '                metadata("Author") = group.name.Substring("META_AUTHOR_".Length)
    '            ElseIf group.name.StartsWith("META_VERSION_") Then
    '                metadata("Version") = group.name.Substring("META_VERSION_".Length)
    '            ElseIf group.name.StartsWith("META_DESC_") Then
    '                metadata("Description") = group.name.Substring("META_DESC_".Length)
    '            End If
    '        Next

    '        Return metadata
    '    End Function

    '#End Region

End Module
