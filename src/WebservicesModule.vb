Imports System.Text
Imports System.Web
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.DataSourcesRaster
Imports System.Net
Imports System.Timers

Public Module WebservicesModule

    Public Function BA_ClipFeatureService(ByVal clipFilePath As String, ByVal webServiceUrl As String, _
                                          ByVal newFilePath As String, ByVal aoiFolder As String) As BA_ReturnCode
        Dim wType As WorkspaceType = BA_GetWorkspaceTypeFromPath(newFilePath)
        If wType = WorkspaceType.Raster Then
            Debug.Print("BA_ClipFeatureService can only write to FileGeodatabase format. Please supply an output path to a FileGeodatabase folder.")
            Return BA_ReturnCode.WriteError
        End If
        Dim sb As StringBuilder = New StringBuilder()
        'url base for query
        sb.Append(webServiceUrl)
        'append trailing backslash to url if it is missing
        If Right(webServiceUrl, 1) <> "/" Then sb.Append("/")
        'append the query; where clause is required; This one returns all records
        Dim whereClause As String = "query?&where={0}"
        sb.Append(String.Format(whereClause, HttpUtility.UrlEncode(String.Format("OBJECTID>{0}", 0))))
        'return all fields
        sb.Append("&outFields=*")
        'return the geometries
        sb.Append("&returnGeometry=true")
        'append the geometry type for spatial query
        sb.Append("&geometryType=esriGeometryEnvelope")
        'append the spatial relation
        sb.Append("&spatialRel=esriSpatialRelIntersects")
        'append the geometry
        sb.Append("&geometry=" & GetJSONEnvelope(clipFilePath))
        'return results in JSON format
        sb.Append("&f=json")
        Dim query As String = sb.ToString
        'read the JSON request
        Dim jsonFeatures As String = GetResult(query)
        Dim byteArray As Byte() = System.Text.Encoding.UTF8.GetBytes(jsonFeatures)
        Dim jsonStream As System.IO.MemoryStream = New System.IO.MemoryStream(byteArray)
        'Check to make sure we got some features, if not IJSONConverterGdb throws an exception
        Dim jsonResult As JsonQueryResult = New JsonQueryResult()
        Dim ser As System.Runtime.Serialization.Json.DataContractJsonSerializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(jsonResult.[GetType]())
        jsonResult = CType(ser.ReadObject(jsonStream), JsonQueryResult)

        If jsonResult IsNot Nothing AndAlso jsonResult.features.Count > 0 Then
            Dim jsonReader As IJSONReader = New JSONReader
            Dim JSONConverterGdb As IJSONConverterGdb = New JSONConverterGdb()
            Dim originalToNewFieldMap As IPropertySet = Nothing
            Dim recordSet As IRecordSet = Nothing
            Dim recordSet2 As IRecordSet2 = Nothing
            Dim workspaceFactory As IWorkspaceFactory = New FileGDBWorkspaceFactory
            Dim workspace As IFeatureWorkspace = Nothing
            Dim searchWS As IWorkspace2 = Nothing
            Dim deleteFClass As IFeatureClass = Nothing
            Dim deleteDataset As IDataset = Nothing
            Try
                jsonReader.ReadFromString(jsonFeatures)
                JSONConverterGdb.ReadRecordSet(jsonReader, Nothing, Nothing, recordSet, originalToNewFieldMap)

                Dim outputFolder As String = "PleaseReturn"
                Dim outputFile As String = BA_GetBareName(newFilePath, outputFolder)
                ' Strip trailing "\" if exists
                If outputFolder(Len(outputFolder) - 1) = "\" Then
                    outputFolder = outputFolder.Remove(Len(outputFolder) - 1, 1)
                End If

                Dim tempFile As String = "webQuery"
                workspace = workspaceFactory.OpenFromFile(outputFolder, 0)
                searchWS = CType(workspace, IWorkspace2)
                recordSet2 = CType(recordSet, IRecordSet2)

                'Delete temp file if it exists
                If searchWS.NameExists(esriDatasetType.esriDTFeatureClass, tempFile) Then
                    deleteFClass = workspace.OpenFeatureClass(tempFile)
                    deleteDataset = CType(deleteFClass, IDataset)
                    deleteDataset.Delete()
                End If

                'Delete output file if it exists
                If searchWS.NameExists(esriDatasetType.esriDTFeatureClass, outputFile) Then
                    deleteFClass = workspace.OpenFeatureClass(outputFile)
                    deleteDataset = CType(deleteFClass, IDataset)
                    deleteDataset.Delete()
                End If

                'Save query results to temp file in target geodatabase
                recordSet2.SaveAsTable(workspace, tempFile)
                'Clip queried layer to aoi
                Dim retVal As Short = BA_ClipAOIVector(aoiFolder, outputFolder & "\" & tempFile, outputFile, outputFolder, True)
                If retVal = 1 Then
                    'Delete temporary query file
                    'Re-initialize workspace to resolve separated RCW error
                    workspace = workspaceFactory.OpenFromFile(outputFolder, 0)
                    deleteFClass = workspace.OpenFeatureClass(tempFile)
                    deleteDataset = CType(deleteFClass, IDataset)
                    deleteDataset.Delete()
                    Return BA_ReturnCode.Success
                End If
                Return BA_ReturnCode.UnknownError
            Catch ex As Exception
                Debug.Print("BA_ClipFeatureService Exception: " & ex.Message)
                Return BA_ReturnCode.UnknownError
            Finally
                recordSet = Nothing
                recordSet2 = Nothing
                workspace = Nothing
                searchWS = Nothing
                deleteFClass = Nothing
                deleteDataset = Nothing
            End Try
        Else
            Return BA_ReturnCode.ReadError
        End If
    End Function

    Private Function GetJSONEnvelope(ByVal clipFilePath As String) As String
        Dim aoiFolder As String = "Please return"
        Dim aoiFile As String = BA_GetBareName(clipFilePath, aoiFolder)
        Dim fClass As IFeatureClass = Nothing
        Dim pClipFCursor As IFeatureCursor = Nothing
        Dim pClipFeature As IFeature = Nothing
        Dim pEnv As IEnvelope = Nothing
        Dim jsonOut As IJSONObject = New JSONObject
        Dim JSONConverter As IJSONConverterGeometry = New JSONConverterGeometry()

        Try
            Dim wType As WorkspaceType = BA_GetWorkspaceTypeFromPath(aoiFolder)
            If wType = WorkspaceType.Geodatabase Then
                fClass = BA_OpenFeatureClassFromGDB(aoiFolder, aoiFile)
            Else
                fClass = BA_OpenFeatureClassFromFile(aoiFolder, aoiFile)
            End If
            If fClass IsNot Nothing Then
                'retrieve IFeature from FeatureClass
                pClipFCursor = fClass.Search(Nothing, False)
                pClipFeature = pClipFCursor.NextFeature
                pEnv = pClipFeature.Shape.Envelope
                'Querying the geometry returned a string that was too long for the url so we use the envelope insead
                JSONConverter.QueryJSONEnvelope(pEnv, False, jsonOut)
                Return jsonOut.ToJSONString(Nothing)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Debug.Print("GetJSONEnvelope Exception: " & ex.Message)
            Return Nothing
        Finally
            fClass = Nothing
            pClipFCursor = Nothing
            pClipFeature = Nothing
            pEnv = Nothing
        End Try
    End Function

    Private Function GetResult(ByVal url As String) As String
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url)
        Dim resp As System.Net.WebResponse = req.GetResponse()
        Using stream As System.IO.Stream = resp.GetResponseStream
            Using streamReader As System.IO.StreamReader = New System.IO.StreamReader(stream)
                Return streamReader.ReadToEnd
            End Using
        End Using
    End Function

    'http://resources.arcgis.com/en/help/arcobjects-net/conceptualhelp/index.html#/How_to_create_an_image_server_layer/00010000047t000000/
    Public Function BA_ClipImageService(ByVal clipFilePath As String, ByVal webServiceUrl As String, _
                                        ByVal newFilePath As String) As BA_ReturnCode

        Dim wType As WorkspaceType = BA_GetWorkspaceTypeFromPath(newFilePath)
        If wType = WorkspaceType.Raster Then
            Debug.Print("BA_ClipImageService can only write to FileGeodatabase format. Please supply an output path to a FileGeodatabase folder.")
            Return BA_ReturnCode.WriteError
        End If

        Dim isLayer As IImageServerLayer = New ImageServerLayerClass
        Dim imageRaster As IRaster = Nothing
        Dim imageRasterProps As IRasterProps = Nothing
        Dim clipRaster As IRaster = Nothing
        Dim clipRasterProps As IRasterProps = Nothing
        Try
            'Create an image server layer by passing a URL.
            Dim URL As String = webServiceUrl
            isLayer.Initialize(URL)
            'Get the raster from the image server layer.
            imageRaster = isLayer.Raster

            'For services that require https/authentication, use the following
            'Dim isLayer As IImageServerLayer = CreateSecuredISLayer("http://server:6080/arcgis/services", "serviceName")

            'The raster from an image server is normally large; Define the size of the raster.
            imageRasterProps = DirectCast(imageRaster, IRasterProps)
            clipRaster = GetClipRaster(clipFilePath)
            clipRasterProps = DirectCast(clipRaster, IRasterProps)

            '@ToDo: May need to worry about the projection in real-life
            imageRasterProps.Extent = clipRasterProps.Extent
            imageRasterProps.Width = clipRasterProps.Width
            imageRasterProps.Height = clipRasterProps.Height

            'Save the clipped raster to the file geodatabase.
            Dim newFolder As String = "PleaseReturn"
            Dim newFile As String = BA_GetBareName(newFilePath, newFolder)
            ' Strip trailing "\" if exists
            If newFolder(Len(newFolder) - 1) = "\" Then
                newFolder = newFolder.Remove(Len(newFolder) - 1, 1)
            End If

            'Remove the target raster if it exists
            Dim retVal As Short = 1
            If BA_File_Exists(newFilePath, WorkspaceType.Geodatabase, esriDatasetType.esriDTRasterDataset) Then
                retVal = BA_RemoveRasterFromGDB(newFolder, newFile)
            End If
            If retVal = 1 Then
                retVal = BA_SaveRasterDatasetGDB(imageRaster, newFolder, BA_RASTER_FORMAT, newFile)
            End If

            If retVal = 1 Then
                Return BA_ReturnCode.Success
            Else
                Return BA_ReturnCode.UnknownError
            End If
        Catch ex As Exception
            Debug.Print("BA_ClipImageService Exception: " & ex.Message)
            Return BA_ReturnCode.UnknownError
        Finally
            imageRaster = Nothing
            imageRasterProps = Nothing
            clipRaster = Nothing
            clipRasterProps = Nothing
        End Try
    End Function

    Private Function GetClipRaster(ByVal clipRasterPath As String) As IRaster
        Dim aoiFolder As String = "PleaseReturn"
        Dim aoiFile As String = BA_GetBareName(clipRasterPath, aoiFolder)
        ' Strip trailing "\" if exists
        If aoiFolder(Len(aoiFolder) - 1) = "\" Then
            aoiFolder = aoiFolder.Remove(Len(aoiFolder) - 1, 1)
        End If
        Dim pGeoDataset As IGeoDataset = Nothing
        Dim pRasterDataset As IRasterDataset = Nothing
        Try
            Dim wType As WorkspaceType = BA_GetWorkspaceTypeFromPath(clipRasterPath)
            If wType = WorkspaceType.Geodatabase Then
                pGeoDataset = BA_OpenRasterFromGDB(aoiFolder, aoiFile)
            Else
                pGeoDataset = BA_OpenRasterFromFile(aoiFolder, aoiFile)
            End If
            pRasterDataset = CType(pGeoDataset, IRasterDataset)
            Return pRasterDataset.CreateDefaultRaster
        Catch ex As Exception
            Debug.Print("GetClipRaster Exception: " & ex.Message)
            Return Nothing
        Finally
            pGeoDataset = Nothing
            pRasterDataset = Nothing
        End Try

    End Function

    'Note that an image service can't be clipped directly to a vector. This function converts a vector to a raster
    'before calling another clip function and deletes the temporary raster when the clip completes
    Public Function BA_ClipImageServiceToVector(ByVal clipFilePath As String, ByVal clipField As String, ByVal webServiceUrl As String, _
                                                ByVal newFilePath As String, ByVal snapRasterPath As String) As BA_ReturnCode
        Dim tempRaster As String = "ClipRas"
        Dim parentPath As String = "PleaseReturn"

        Try
            Dim cellSize As Double = BA_CellSizeImageService(webServiceUrl)
            Dim clipFeatures As String = BA_GetBareName(clipFilePath, parentPath)
            Dim outRasterPath As String = parentPath & tempRaster
            Dim success As BA_ReturnCode = BA_Feature2RasterGP(clipFilePath, outRasterPath, clipField, cellSize, snapRasterPath)
            If success = BA_ReturnCode.Success Then
                BA_ClipImageService(outRasterPath, webServiceUrl, newFilePath)
            End If
            Return success
        Catch ex As Exception
            Debug.Print("BA_ClipImageServiceToVector Exception: " & ex.Message)
            Return BA_ReturnCode.UnknownError
        Finally
            Dim wType = BA_GetWorkspaceTypeFromPath(parentPath)
            If wType = WorkspaceType.Geodatabase Then
                BA_RemoveRasterFromGDB(parentPath, tempRaster)
            Else
                BA_Remove_Raster(parentPath, tempRaster)
            End If
        End Try
    End Function

    Public Function BA_QueryFeatureServiceFieldNames(ByVal webserviceUrl As String, ByVal fieldType As esriFieldType) As IList(Of String)
        Dim sb As StringBuilder = New StringBuilder()
        'read the JSON request
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(webserviceUrl & "?f=pjson")
        Dim resp As System.Net.WebResponse = req.GetResponse()

        Dim fs As FeatureService = New FeatureService()
        Dim ser As System.Runtime.Serialization.Json.DataContractJsonSerializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(fs.[GetType]())
        fs = CType(ser.ReadObject(resp.GetResponseStream), FeatureService)
        Dim fieldList As List(Of String) = New List(Of String)
        For Each fsf As FeatureServiceField In fs.fields
            If fsf.fieldType = fieldType Then
                fieldList.Add(fsf.alias)
            End If
        Next
        If fieldList.Count > 1 Then
            fieldList.Sort()
        End If
        Return fieldList
    End Function

    'Public Function BA_UploadAoi(ByVal webserviceUrl As String, ByVal strToken As String) As BA_ReturnCode
    '    Dim reqT As HttpWebRequest
    '    Dim resT As HttpWebResponse
    '    'The end point for getting a token for the web service
    '    reqT = WebRequest.Create(webserviceUrl)
    '    'This is a POST request
    '    reqT.Method = "POST"
    '    'We are sending a form
    '    reqT.ContentType = "multipart/form-data"

    '    'Retrieve the token and format it for the header; Token comes from caller
    '    Dim cred As String = String.Format("{0} {1}", "Token", strToken)
    '    'Put token in header
    '    reqT.Headers(HttpRequestHeader.Authorization) = cred

    '    'These are the field/value pairs that would be on an html form
    '    Dim Data As String = "filename=aoi_text&md5=eae7186b59e33aa610ac29a4dc952caa&file="
    '    'Dim filePath As String = "C:\Docs\Lesley\aoi1_05222013.zip"
    '    Dim filePath As String = "C:\Docs\Lesley\Landis\data\TRY2\Landis-log.txt"
    '    Dim fileInfo As System.IO.FileInfo = New System.IO.FileInfo(filePath)
    '    'Encode them to Byte format to include with the request
    '    Dim dataArray As Byte() = Encoding.UTF8.GetBytes(Data)
    '    reqT.ContentLength = dataArray.Length + fileInfo.Length

    '    Try
    '        'Intercept the httpRequest so we can add the request fields
    '        Dim dataStreamT As System.IO.Stream = reqT.GetRequestStream()
    '        dataStreamT.Write(dataArray, 0, dataArray.Length)

    '        ' Read the source file into a byte array.
    '        Dim bytes() As Byte = New Byte((fileInfo.Length) - 1) {}
    '        Dim numBytesToRead As Integer = CType(fileInfo.Length, Integer)
    '        Dim numBytesRead As Integer = 0

    '        Using fsSource As System.IO.FileStream = New System.IO.FileStream(filePath, System.IO.FileMode.Open, _
    '                                                                          System.IO.FileAccess.Read)
    '            While (numBytesToRead > 0)
    '                ' Read may return anything from 0 to numBytesToRead.
    '                Dim n As Integer = fsSource.Read(bytes, numBytesRead, numBytesToRead)
    '                ' Break when the end of the file is reached.
    '                If (n = 0) Then
    '                    Exit While
    '                End If
    '                numBytesRead = (numBytesRead + n)
    '                numBytesToRead = (numBytesToRead - n)
    '            End While
    '            numBytesToRead = bytes.Length

    '            'write data to request
    '            dataStreamT.Write(bytes, 0, numBytesToRead)
    '        End Using
    '        dataStreamT.Close()

    '        resT = CType(reqT.GetResponse(), HttpWebResponse)
    '        'Printing the response to the Console for testing
    '        Using SReader As System.IO.StreamReader = New System.IO.StreamReader(resT.GetResponseStream)
    '            Debug.Print(SReader.ReadToEnd())
    '        End Using
    '        'If we didn't get an exception, the upload was successful
    '        Return BA_ReturnCode.Success
    '    Catch ex As WebException
    '        Debug.Print("BA_UploadAoi Exception: " & ex.Message)
    '        Return BA_ReturnCode.UnknownError
    '    End Try

    'End Function

    Public Function BA_UploadMultiPart(ByVal webserviceUrl As String, ByVal strToken As String, _
                                       ByVal fileName As String, ByVal filePath As String, _
                                       ByVal comment As String) As AoiTask
        Dim reqT As HttpWebRequest
        Dim resT As HttpWebResponse
        Dim anUpload As AoiTask = New AoiTask
        'The end point for getting a token for the web service
        reqT = WebRequest.Create(webserviceUrl)
        'This is a POST request
        reqT.Method = "POST"
        'We are sending a form
        Dim boundary As String = MultipartFormHelper.CreateFormDataBoundary()
        reqT.ContentType = "multipart/form-data; boundary=" & boundary
        reqT.KeepAlive = True

        'Retrieve the token and format it for the header; Token comes from caller
        Dim cred As String = String.Format("{0} {1}", "Token", strToken)
        'Put token in header
        reqT.Headers(HttpRequestHeader.Authorization) = cred

        Try
            Dim requestStream As System.IO.Stream = reqT.GetRequestStream
            Dim postData As Dictionary(Of String, String) = New Dictionary(Of String, String)
            postData.Add("filename", fileName)
            If Not String.IsNullOrEmpty(comment) Then postData.Add("comment", Trim(comment))

            Dim fileInfo As System.IO.FileInfo = New System.IO.FileInfo(filePath)
            postData.Add("md5", MultipartFormHelper.GenerateMD5Hash(fileInfo))
            MultipartFormHelper.WriteMultipartFormData(postData, requestStream, boundary)

            If fileInfo IsNot Nothing Then
                '@ToDo: Remove hard-coding; write a dynamic function to determine mime type
                'Dim fileMimeType As String = "text/plain"
                Dim fileMimeType As String = BA_Mime_Zip
                Dim fileFormKey As String = "file"
                MultipartFormHelper.WriteMultipartFormData(fileInfo, requestStream, boundary, fileMimeType, fileFormKey)
            End If
            Dim endBytes() As Byte = Encoding.UTF8.GetBytes("--" + boundary + "--")
            requestStream.Write(endBytes, 0, endBytes.Length)
            requestStream.Close()

            resT = CType(reqT.GetResponse(), HttpWebResponse)
            'Convert the JSON response to a Task object
            Dim ser As System.Runtime.Serialization.Json.DataContractJsonSerializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(anUpload.[GetType]())
            'Put JSON payload into AOI object
            anUpload = CType(ser.ReadObject(resT.GetResponseStream), AoiTask)

            'If we didn't get an exception, the upload was successful
            Return anUpload
        Catch w As WebException
            Dim exceptResp As HttpWebResponse = TryCast(w.Response, HttpWebResponse)
            Dim sb As StringBuilder = New StringBuilder
            'The response is a long html page
            'The exception is indicated with this line: <pre class="exception_value">An AOI of the same name already exists.</pre>
            '@ToDo: Figure out how to parse the response and pull out this exception_value
            If exceptResp IsNot Nothing Then
                Using SReader As System.IO.StreamReader = New System.IO.StreamReader(exceptResp.GetResponseStream)
                    sb.Append(SReader.ReadToEnd)
                End Using
            End If
            Debug.Print("BA_UploadMultiPart WebException: " & sb.ToString)
            Return anUpload
        Catch ex As Exception
            Debug.Print("BA_UploadMultiPart: " & ex.Message)
            Return anUpload
        End Try

    End Function

    Private Sub DebugPropertySet(ByVal propertySet As IPropertySet)
        Dim names(propertySet.Count - 1) As Object
        Dim values(propertySet.Count - 1) As Object
        propertySet.GetAllProperties(names, values)
        For i As Integer = 0 To propertySet.Count - 1
            Debug.Print(CStr(names(i) & "-->"))
            Debug.Print(values(i).ToString & vbCrLf)
        Next
    End Sub

    Public Function BA_List_Aoi(ByVal url As String, ByVal strToken As String) As Dictionary(Of String, StoredAoi)
        Dim aoiDictionary As Dictionary(Of String, StoredAoi) = New Dictionary(Of String, StoredAoi)

        Dim reqT As HttpWebRequest
        Dim resT As HttpWebResponse
        'The end point for getting a token for the web service
        url = url & "aois/"
        reqT = WebRequest.Create(url)
        'This is a GET request
        reqT.Method = "GET"

        'Retrieve the token and format it for the header; Token comes from caller
        Dim cred As String = String.Format("{0} {1}", "Token", strToken)
        'Put token in header
        reqT.Headers(HttpRequestHeader.Authorization) = cred

        Try
            resT = CType(reqT.GetResponse(), HttpWebResponse)

            'Convert the JSON response to StoredAoiArray object
            Dim storedAois As StoredAoiArray = New StoredAoiArray
            Dim ser As System.Runtime.Serialization.Json.DataContractJsonSerializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(storedAois.[GetType]())
            storedAois = CType(ser.ReadObject(resT.GetResponseStream), StoredAoiArray)
            'Populate Dictionary from storedAois array
            For Each sAoi As StoredAoi In storedAois.results
                aoiDictionary.Add(sAoi.name, sAoi)
            Next
        Catch ex As WebException
            Debug.Print("BA_List_Aoi Exception: " & ex.Message)
        End Try
        Return aoiDictionary
    End Function

    Public Function BA_Download_Aoi(ByVal url As String, ByVal strToken As String) As AoiTask
        Dim reqT As HttpWebRequest
        Dim resT As HttpWebResponse
        Dim aDownload As AoiTask = New AoiTask
        'The end point for getting a token for the web service
        reqT = WebRequest.Create(url)
        'This is a GET request
        reqT.Method = "GET"

        'Retrieve the token and format it for the header; Token comes from caller
        Dim cred As String = String.Format("{0} {1}", "Token", strToken)
        'Put token in header
        reqT.Headers(HttpRequestHeader.Authorization) = cred
        Try
            resT = CType(reqT.GetResponse(), HttpWebResponse)
            'Convert the JSON response to an AoiUpload object
            Dim ser As System.Runtime.Serialization.Json.DataContractJsonSerializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(aDownload.[GetType]())
            'Put JSON payload into AOI object
            aDownload = CType(ser.ReadObject(resT.GetResponseStream), AoiTask)

            'Using stream As System.IO.Stream = resT.GetResponseStream
            '    Using streamReader As System.IO.StreamReader = New System.IO.StreamReader(stream)
            '        Debug.Print("JSON response -->" & streamReader.ReadToEnd)
            '    End Using
            'End Using

            'If we didn't get an exception, the upload was successful
            Return aDownload
        Catch ex As WebException
            Debug.Print("BA_Download_Aoi Exception: " & ex.Message)
            Return aDownload
        End Try
    End Function

    Public Function BA_GetResponseContentType(ByVal url As String, ByVal token As String) As String
        Dim reqT As HttpWebRequest
        Dim resT As HttpWebResponse
        Try
            reqT = WebRequest.Create(url)
            'This is a GET request
            reqT.Method = "HEAD"

            'Retrieve the token and format it for the header; Token comes from caller
            Dim cred As String = String.Format("{0} {1}", "Token", token)
            'Put token in header
            reqT.Headers(HttpRequestHeader.Authorization) = cred
            resT = CType(reqT.GetResponse(), HttpWebResponse)
            Return resT.ContentType
        Catch ex As WebException
            Debug.Print("TestDownload: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function BA_DownloadFile(ByVal url As String, ByVal token As String, ByVal filePath As String) As BA_ReturnCode
        ' Using WebClient for built-in file download functionality
        Dim myWebClient As New WebClient()
        Try
            'Retrieve the token and format it for the header; Token comes from caller
            Dim cred As String = String.Format("{0} {1}", "Token", token)
            'Put token in header
            myWebClient.Headers(HttpRequestHeader.Authorization) = cred
            myWebClient.DownloadFile(url, filePath)
            Return BA_ReturnCode.Success
        Catch ex As Exception
            Debug.Print("BA_DownloadFile: " & ex.Message)
            Return BA_ReturnCode.UnknownError
        End Try
    End Function

    Public Function BA_AoiInArchive(ByVal url As String, ByVal strToken As String, _
                                    ByVal aoiName As String) As Boolean
        Dim storedAois As Dictionary(Of String, StoredAoi) = BA_List_Aoi(url, strToken)
        For Each kvp As KeyValuePair(Of String, StoredAoi) In storedAois
            If kvp.Value.name.ToUpper = aoiName.ToUpper Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function BA_File_ExistsImageServer(ByVal imageUrl As String) As Boolean
        Dim isLayer As IImageServerLayer = New ImageServerLayerClass
        Dim imageRaster As IRaster = Nothing
        Try
            isLayer.Initialize(imageUrl)
            imageRaster = isLayer.Raster
            Return True
        Catch ex As Exception
            ' An exception was thrown while trying to open the dataset, return false
            Return False
        End Try
    End Function

    ' Calculate raster cellsize from an input GeoDataset
    Public Function BA_CellSizeImageService(ByVal url As String) As Double
        Dim isLayer As IImageServerLayer = New ImageServerLayerClass
        Dim imageRaster As IRaster = Nothing
        Dim imageRasterProps As IRasterProps = Nothing
        Try
            'Create an image server layer by passing a URL.
            isLayer.Initialize(url)
            'Get the raster from the image server layer.
            imageRaster = isLayer.Raster
            imageRasterProps = DirectCast(imageRaster, IRasterProps)
            Dim pPnt As IPnt = imageRasterProps.MeanCellSize
            Return (pPnt.X + pPnt.Y) / 2
        Catch ex As Exception
            MsgBox("BA_CellSizeImageService: " & ex.Message)
            Return 0
        Finally
            isLayer = Nothing
            imageRaster = Nothing
            imageRasterProps = Nothing
        End Try
    End Function

    Public Function BA_FeatureServiceSpatialReference(ByVal url As String) As ISpatialReference
        Dim sb As StringBuilder = New StringBuilder()
        Dim spRef As ISpatialReference = Nothing
        'read the JSON request
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url & "?f=pjson")
        Dim resp As System.Net.WebResponse = req.GetResponse()

        Dim fs As FeatureService = New FeatureService()
        Dim ser As System.Runtime.Serialization.Json.DataContractJsonSerializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(fs.[GetType]())
        fs = CType(ser.ReadObject(resp.GetResponseStream), FeatureService)

        If fs.extent.spatialReference IsNot Nothing Then
            Dim factory As ISpatialReferenceFactory3 = New SpatialReferenceEnvironment()
            spRef = factory.CreateSpatialReference(fs.extent.spatialReference.wkid)
        End If
        Return spRef
    End Function

    Public Function BA_File_ExistsFeatureServer(ByVal url As String) As Boolean
        Dim sb As StringBuilder = New StringBuilder()
        Try

            'read the JSON request
            Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(url & "?f=pjson")
            Dim resp As System.Net.WebResponse = req.GetResponse()

            Dim fs As FeatureService = New FeatureService()
            Dim ser As System.Runtime.Serialization.Json.DataContractJsonSerializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(fs.[GetType]())
            fs = CType(ser.ReadObject(resp.GetResponseStream), FeatureService)
            If fs IsNot Nothing Then
                If Not String.IsNullOrEmpty(fs.name) Then
                    Return True
                End If
            End If
            Return False
        Catch ex As Exception
            'We encountered an exception; Feature service doesn't exist
            Return False
        End Try
    End Function

    Public Function BA_QueryAllFeatureServiceFieldNames(ByVal webserviceUrl As String) As IList(Of FeatureServiceField)
        Dim sb As StringBuilder = New StringBuilder()
        Dim fieldList As List(Of FeatureServiceField) = New List(Of FeatureServiceField)
        'read the JSON request
        Dim req As System.Net.WebRequest = System.Net.WebRequest.Create(webserviceUrl & "?f=pjson")
        Dim resp As System.Net.WebResponse = req.GetResponse()

        Dim fs As FeatureService = New FeatureService()
        Dim ser As System.Runtime.Serialization.Json.DataContractJsonSerializer = New System.Runtime.Serialization.Json.DataContractJsonSerializer(fs.[GetType]())
        fs = CType(ser.ReadObject(resp.GetResponseStream), FeatureService)
        For Each fsf As FeatureServiceField In fs.fields
            fieldList.Add(fsf)
        Next
        Return fieldList
    End Function
End Module
