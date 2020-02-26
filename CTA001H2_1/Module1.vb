Imports System.Data
Imports Teradata.Client.Provider
Imports System.IO
Imports System.Configuration
Imports System.IO.Compression
Imports System.Globalization
Module Module1

    Sub Main(ByVal args() As String)
        Try
            Console.WriteLine("Start Program: " & GetAppKey("PROGRAM_NAME") & vbCrLf)
            'Create Data Set
            Dim myQuery As String = QueryBuilder(args, GetAppKey("QUERY"))
            If GetAppKey("PRINT_QUERY") = "Y" Then
                Console.WriteLine("")
                Console.WriteLine(myQuery)
                Console.WriteLine("")
            End If
            Dim MyDataSet = New DataSet
            MyDataSet = GetDataSetTd(myQuery)

            Dim myQuery2 As String = QueryBuilder(args, GetAppKey("QUERY2"))
            If GetAppKey("PRINT_QUERY") = "Y" Then
                Console.WriteLine("")
                Console.WriteLine(myQuery)
                Console.WriteLine("")
            End If
            Dim MyDataSet2 = New DataSet
            MyDataSet2 = GetDataSetTd(myQuery2)

            Dim myQuery3 As String = QueryBuilder(args, GetAppKey("QUERY3"))
            If GetAppKey("PRINT_QUERY") = "Y" Then
                Console.WriteLine("")
                Console.WriteLine(myQuery)
                Console.WriteLine("")
            End If
            Dim MyDataSet3 = New DataSet
            MyDataSet3 = GetDataSetTd(myQuery3)

            'Create Text File
            Dim myFileName As String = CreateFileName(args, GetAppKey("FILENAME"))
            CreateFile(MyDataSet, MyDataSet2, MyDataSet3, GetAppKey("DIRECTORY") & myFileName)

            'Create Zip File
            CreateZipFile(GetAppKey("DIRECTORY") & myFileName)

            'Run FTP 
            If GetAppKey("RUN_FTP") = "Y" Then
                If GetAppKey("FTP_ZIP") = "Y" Then
                    myFileName = myFileName & ".gz"
                Else
                    myFileName = myFileName
                End If
                FTPSent(GetAppKey("FTP") & myFileName, GetAppKey("DIRECTORY") & myFileName, GetAppKey("USERNAME"), GetAppKey("PASSWORD"))
            End If

            Console.WriteLine("")
            Console.WriteLine("Done...")

        Catch ex As Exception
            Console.WriteLine("Error Main Module: " & ex.ToString)
        Finally

            If GetAppKey("READ_KEY") = "Y" Then
                Console.ReadKey()
            End If
        End Try
    End Sub
    Function QueryBuilder(ByVal args() As String, ByVal MyQuery As String) As String
        Try
            Dim str As String = MyQuery
            If args.Count >= 1 Then
                For i As Integer = 0 To args.Length - 1
                    'Console.WriteLine("@ARGS" & (i + 1).ToString)
                    str = str.Replace("@ARGS" & (i + 1).ToString, args(i))
                    'Console.WriteLine(str)
                Next
                QueryBuilder = str
            Else
                QueryBuilder = MyQuery
            End If
        Catch ex As Exception
            Console.WriteLine("QueryBuilder error: " & ex.ToString)
            QueryBuilder = MyQuery
        End Try
    End Function
    Function GetDataSetTd(ByVal myQuery As String) As DataSet
        'Set up connection string
        Dim myText As String = "Connecting to database using teradata" & vbCrLf & vbCrLf
        Dim connectionString As String = GetAppKey("CONN_STR")

        Dim conn As TdConnection = New TdConnection(connectionString)
        Try
            conn.Open()
            Console.WriteLine(myText & "Connection opened, Process Dataset" & vbCrLf)
            'Console.WriteLine(myQuery)

            Dim tout As Integer = CInt(GetAppKey("TIMEOUT"))
            Dim command = New TdCommand(myQuery, conn)
            command.CommandTimeout = tout
            Dim adapter = New TdDataAdapter(command)

            Dim MyDataSet = New DataSet
            adapter.Fill(MyDataSet)
            GetDataSetTd = MyDataSet

        Catch ex As TdException
            Console.WriteLine(myText & "Error: " & ex.ToString & vbCrLf)
            GetDataSetTd = New DataSet
        Finally
            ' Close connection
            conn.Close()
            Console.WriteLine(myText & "Connection closed." & vbCrLf)
        End Try
    End Function
    Function CreateFileName(ByVal args() As String, ByVal MyFileName As String) As String
        Try
            Dim str As String = MyFileName
            For i As Integer = 0 To args.Length - 1
                str = str.Replace("@ARGS" & (i + 1).ToString, args(i))
            Next
            CreateFileName = str
        Catch ex As Exception
            Console.WriteLine("CreateFileName error: " & ex.ToString)
            CreateFileName = MyFileName
        End Try
    End Function
    Sub CreateFile(ByVal MyDataSet As DataSet, ByVal MyDataSet2 As DataSet, ByVal MyDataSet3 As DataSet, ByVal MyFileName As String)
        Try
            Dim MyDsTable = MyDataSet.Tables(0)
            Dim columnCount = MyDsTable.Columns.Count

            Dim MyDsTable2 = MyDataSet2.Tables(0)
            Dim columnCount2 = MyDsTable2.Columns.Count

            Dim MyDsTable3 = MyDataSet3.Tables(0)
            Dim columnCount3 = MyDsTable3.Columns.Count
            Console.WriteLine("Create Textfile: " & MyFileName & vbCrLf)
            'Console.WriteLine("Column Count: " & columnCount)

            Dim row As DataRow
            Console.WriteLine()
            Dim mystr As String = ""
            Dim counter As Integer
            Using sw As StreamWriter = New StreamWriter(MyFileName)
                counter = 1

                Dim strDlm As String = GetAppKey("DELIMITER")
                Dim iCabang As String = ""
                Dim iCabang2 As String = ""
                Dim Strip As String = ""
                Dim iAccTypeSubcat As String = ""
                Dim FullStrip As String = ""
                Dim iPage As Integer = 1
                Dim iNomor As Integer = 0
                Dim Pemisah As Integer = 0
                Dim Pemisah2 As Integer = 0
                Dim IDX As Integer = 0
                Dim IDX2 As Integer = 0
                Dim Cont As Integer = 0
                Dim HeadPage As Boolean = True
                Dim AccType As Boolean = True
                Dim First As Boolean = True
                Dim First2 As Boolean = True
                Dim First3 As Boolean = True
                Dim First4 As Boolean = True
                Dim First5 As Boolean = True
                Dim First6 As Boolean = True
                Dim TOT_PLAFON As Decimal
                Dim TOT_OUTSTANDING As Decimal
                Dim current As CultureInfo = New System.Globalization.CultureInfo("id-ID")
                'Console.WriteLine("The current UI culture is {0}", current.Name)
                FullStrip = "-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

                For Each row In MyDsTable.Rows
                    If First5 = True Then
                        iCabang2 = CStr(row("BRANCH_NAME"))
                        iCabang = iCabang2
                        First5 = False
                    Else
                        iCabang = CStr(row("BRANCH_NAME"))
                    End If
                    If iCabang <> "" Then
ATAS:
                        Cont = CStr(row("CONT"))
                        Pemisah = CInt(row("IDX"))
                        If (Pemisah = 1 Or Pemisah = 4) Then
                            Strip = StringBuilder(96, "", "------------------------------------------------")
                        ElseIf (Pemisah = 2 Or Pemisah = 3 Or Pemisah = 5 Or Pemisah = 6) Then
                            Strip = StringBuilder(96, "", "------------------------------------------------")
                        End If

                        If HeadPage = True Then
                            If iPage <> 1 Then

                                mystr = Chr(12) & "PT BANK NEGARA INDONESIA (Persero) Tbk                                                                                                                                                                                                           CTA001H2"
                                sw.WriteLine(mystr)
                                IDX = IDX + 1
                            Else
                                mystr = "PT BANK NEGARA INDONESIA (Persero) Tbk                                                                                                                                                                                                           CTA001H2"
                                sw.WriteLine(mystr)
                                IDX = IDX + 1
                            End If
                            mystr = StringBuilder(1, "", "Cabang : " & iCabang2)
                            mystr = StringBuilder(242, mystr, row("TANGGAL"))
                            sw.WriteLine(mystr)
                            IDX = IDX + 1
                            If iPage <= 999 Then
                                mystr = StringBuilder(242, "", "HAL. " & CStr(iPage))
                                IDX = IDX + 1
                            Else
                                mystr = StringBuilder(242, "", "HAL. ***")
                                IDX = IDX + 1
                            End If
                            sw.WriteLine(mystr)
                            IDX = IDX + 1
                            sw.WriteLine("")
                            IDX = IDX + 1
                            mystr = StringBuilder(118, "", "DAFTAR GARANSI BANK YG DITERBITKAN")
                            sw.WriteLine(mystr)
                            IDX = IDX + 1
                            sw.WriteLine("")
                            IDX = IDX + 1
                            mystr = StringBuilder(118, "", "Tanggal : " & row("TANGGAL"))
                            sw.WriteLine(mystr)
                            IDX = IDX + 1
                            sw.WriteLine("")
                            sw.WriteLine("")
                            IDX = IDX + 2
                            mystr = FullStrip & vbCrLf
                            mystr = mystr & "NO.   NO. CIF            NO. REKENING     NAMA NASABAH                   KODE KLASI  CCY  SEG           PERSETUJUAN          OUTSTANDING          TGL BUKA   TGL TUTUP    TGL KLAIM  TIPE" & vbCrLf
                            mystr = mystr & "                                                                         FIKASI GL" & vbCrLf
                            mystr = mystr & FullStrip
                            sw.WriteLine(mystr)
                            IDX = IDX + 4
                            HeadPage = False
                        End If

                        If First4 = False Then
                            If iCabang <> iCabang2 Then
                                If IDX <> 0 Then
                                    iPage = iPage + 1
                                End If
                                sw.WriteLine(FullStrip)
                                mystr = Chr(12) & "PT BANK NEGARA INDONESIA (Persero) Tbk                                                                                                                                                                                                           CTA001H2"
                                sw.WriteLine(mystr)

                                mystr = StringBuilder(1, "", "Cabang : " & iCabang2)
                                mystr = StringBuilder(242, mystr, row("TANGGAL"))
                                sw.WriteLine(mystr)
                                If iPage <= 999 Then
                                    mystr = StringBuilder(242, "", "HAL. " & CStr(iPage))
                                Else
                                    mystr = StringBuilder(242, "", "HAL. ***")
                                End If
                                sw.WriteLine(mystr)
                                sw.WriteLine("")
                                mystr = StringBuilder(118, "", "DAFTAR GARANSI BANK YG DITERBITKAN")
                                sw.WriteLine(mystr)
                                sw.WriteLine("")
                                mystr = StringBuilder(118, "", "Tanggal :  " & row("TANGGAL"))
                                sw.WriteLine(mystr)
                                sw.WriteLine("")
                                sw.WriteLine("")
                                mystr = FullStrip & vbCrLf
                                mystr = mystr & "REKAPITULASI             JUMLAH REKENING                                TOTAL OUTSTANDING" & vbCrLf
                                mystr = mystr & "" & vbCrLf
                                mystr = mystr & FullStrip
                                sw.WriteLine(mystr)
                                For Each row3 In MyDsTable3.Rows
                                    mystr = ""
                                    If iCabang2 = CStr(row3("BRANCH_NAME")) Then

                                        mystr = StringBuilder(7, "", row3("BI_CCY"))
                                        mystr = StringBuilder(31, mystr, CInt(row3("CONT")))
                                        mystr = StringBuilder(36, mystr, "Rekening")
                                        mystr = StringBuilder(75, mystr, CDec(row3("REKAP_OUTSTANDING")).ToString("#,#0.00", current).PadLeft(20, " "))
                                        sw.WriteLine(mystr)

                                    End If
                                Next

                                sw.WriteLine(FullStrip)

                                mystr = Chr(12) & "PT BANK NEGARA INDONESIA (Persero) Tbk                                                                                                                                                                                                           CTA001H2"
                                sw.WriteLine(mystr)

                                mystr = StringBuilder(1, "", "Cabang : " & iCabang2)
                                mystr = StringBuilder(242, mystr, row("TANGGAL"))
                                sw.WriteLine(mystr)
                                If iPage <= 999 Then
                                    mystr = StringBuilder(242, "", "HAL. " & CStr(iPage))
                                Else
                                    mystr = StringBuilder(242, "", "HAL. ***")
                                End If
                                sw.WriteLine(mystr)
                                sw.WriteLine("")
                                mystr = StringBuilder(118, "", "DAFTAR GARANSI BANK YG DITERBITKAN")
                                sw.WriteLine(mystr)
                                sw.WriteLine("")
                                mystr = StringBuilder(118, "", "Tanggal :  " & row("TANGGAL"))
                                sw.WriteLine(mystr)
                                sw.WriteLine("")
                                sw.WriteLine("")
                                mystr = FullStrip & vbCrLf
                                mystr = mystr & "KODE KLASIFIKASI GL                                                     TOTAL OUTSTANDING" & vbCrLf
                                mystr = mystr & "" & vbCrLf
                                mystr = mystr & FullStrip
                                sw.WriteLine(mystr)


                                For Each row2 In MyDsTable2.Rows
                                    mystr = ""
                                    If iCabang2 = CStr(row2("BRANCH_NAME")) Then

                                        mystr = StringBuilder(7, "", row2("GL_ACCOUNT_ID"))
                                        mystr = StringBuilder(75, mystr, CDec(row2("GL_OUTSTANDING")).ToString("#,#0.00", current).PadLeft(20, " "))


                                        sw.WriteLine(mystr)

                                    End If

                                Next
                                sw.WriteLine(FullStrip)
                                iNomor = 0
                                IDX = 0
                                HeadPage = True
                                Cont = 0
                                iCabang2 = iCabang
                                GoTo ATAS

                            End If
                        End If
                        First4 = False

                        TOT_PLAFON = CDec(row("SUM_PLAFON"))
                        TOT_OUTSTANDING = CDec(row("SUM_OUTSTANDING"))


                        If (Pemisah = 1 Or Pemisah = 4) Then

                            If AccType = True Then
                                mystr = " " & StringBuilder(1, "", CStr(row("ACCOUNT_TYPE")) & "  " & CStr(row("SUB_CATEGORY")) & " " & "-" & " " & row("PRD_NAME_CTA"))
                                sw.WriteLine(mystr)
                                IDX = IDX + 1
                                AccType = False
                            End If

                            mystr = StringBuilder(1, "", CStr(iNomor + 1) & ".")
                            mystr = StringBuilder(7, mystr, row("CIF_KEY").ToString.PadLeft(10, "0"))
                            mystr = StringBuilder(26, mystr, row("NO_REKENING").ToString.PadLeft(16, "0"))
                            mystr = StringBuilder(43, mystr, row("NAMA_NASABAH"))
                            mystr = StringBuilder(76, mystr, row("GL_ACCOUNT_ID"))
                            mystr = StringBuilder(86, mystr, row("BI_CCY"))
                            mystr = StringBuilder(91, mystr, row("SEGMEN"))
                            mystr = StringBuilder(102, mystr, CDec(row("PLAFON")).ToString("#,#.00", current).PadLeft(17, " "))
                            mystr = StringBuilder(126, mystr, CDec(row("OUTSTANDING")).ToString("#,#.00", current).PadLeft(17, " "))
                            mystr = StringBuilder(145, mystr, row("TGL_BUKA"))
                            mystr = StringBuilder(158, mystr, row("TGL_TUTUP"))
                            mystr = StringBuilder(171, mystr, row("TGL_KLAIM"))
                            mystr = StringBuilder(182, mystr, row("TIPE"))
                            sw.WriteLine(mystr)
                            IDX = IDX + 1
                            iNomor = iNomor + 1
                            First2 = True
                        End If


                        If Cont = iNomor Then
                            sw.WriteLine(Strip)
                            IDX = IDX + 1
                            mystr = StringBuilder(7, "", "Sub total")
                            mystr = StringBuilder(105, mystr, TOT_PLAFON.ToString("#,#0.00", current).PadLeft(17, " "))
                            mystr = StringBuilder(126, mystr, TOT_OUTSTANDING.ToString("#,#0.00", current).PadLeft(17, " "))
                            sw.WriteLine(mystr)
                            IDX = IDX + 1
                            sw.WriteLine(Strip)
                            IDX = IDX + 1
                            sw.WriteLine("")
                            IDX = IDX + 1
                            iNomor = 0
                            AccType = True
                            First6 = False
                        End If

                        If IDX > 47 Then
                            IDX = 0
                            If First = True Then
                                If First6 = True Then
                                    sw.WriteLine("")
                                End If
                                sw.WriteLine(FullStrip)
                                First = False
                            Else
                                If First6 = True Then
                                    sw.WriteLine("")
                                End If
                                sw.WriteLine(FullStrip)
                            End If
                            iPage = iPage + 1
                            HeadPage = True
                        End If
                        First6 = True
                        'DEBUG
                        If iPage = 29 Then
                            iPage = iPage
                        End If
                        'DEBUG

                    End If
                    
                    iCabang2 = iCabang
                Next

                sw.WriteLine("")
                sw.WriteLine(FullStrip)

                If iCabang = "YOGYAKARTA" Then
                    mystr = Chr(12) & "PT BANK NEGARA INDONESIA (Persero) Tbk                                                                                                                                                                                                           CTA001H2"
                    sw.WriteLine(mystr)

                    mystr = StringBuilder(1, "", "Cabang : " & iCabang2)
                    mystr = StringBuilder(242, mystr, row("TANGGAL"))
                    sw.WriteLine(mystr)
                    If iPage <= 999 Then
                        mystr = StringBuilder(242, "", "HAL. " & CStr(iPage))
                    Else
                        mystr = StringBuilder(242, "", "HAL. ***")
                    End If
                    sw.WriteLine(mystr)
                    sw.WriteLine("")
                    mystr = StringBuilder(118, "", "DAFTAR GARANSI BANK YG DITERBITKAN")
                    sw.WriteLine(mystr)
                    sw.WriteLine("")
                    mystr = StringBuilder(118, "", "Tanggal :  " & row("TANGGAL"))
                    sw.WriteLine(mystr)
                    sw.WriteLine("")
                    sw.WriteLine("")
                    mystr = FullStrip & vbCrLf
                    mystr = mystr & "REKAPITULASI             JUMLAH REKENING                                TOTAL OUTSTANDING" & vbCrLf
                    mystr = mystr & "" & vbCrLf
                    mystr = mystr & FullStrip
                    sw.WriteLine(mystr)
                    For Each row3 In MyDsTable3.Rows
                        mystr = ""
                        If iCabang2 = CStr(row3("BRANCH_NAME")) Then

                            mystr = StringBuilder(7, "", row3("BI_CCY"))
                            mystr = StringBuilder(31, mystr, CInt(row3("CONT")))
                            mystr = StringBuilder(36, mystr, "Rekening")
                            mystr = StringBuilder(75, mystr, CDec(row3("REKAP_OUTSTANDING")).ToString("#,#0.00", current).PadLeft(20, " "))
                            sw.WriteLine(mystr)

                        End If
                    Next

                    sw.WriteLine(FullStrip)

                    mystr = Chr(12) & "PT BANK NEGARA INDONESIA (Persero) Tbk                                                                                                                                                                                                           CTA001H2"
                    sw.WriteLine(mystr)

                    mystr = StringBuilder(1, "", "Cabang : " & iCabang2)
                    mystr = StringBuilder(242, mystr, row("TANGGAL"))
                    sw.WriteLine(mystr)
                    If iPage <= 999 Then
                        mystr = StringBuilder(242, "", "HAL. " & CStr(iPage))
                    Else
                        mystr = StringBuilder(242, "", "HAL. ***")
                    End If
                    sw.WriteLine(mystr)
                    sw.WriteLine("")
                    mystr = StringBuilder(118, "", "DAFTAR GARANSI BANK YG DITERBITKAN")
                    sw.WriteLine(mystr)
                    sw.WriteLine("")
                    mystr = StringBuilder(118, "", "Tanggal :  " & row("TANGGAL"))
                    sw.WriteLine(mystr)
                    sw.WriteLine("")
                    sw.WriteLine("")
                    mystr = FullStrip & vbCrLf
                    mystr = mystr & "KODE KLASIFIKASI GL                                                     TOTAL OUTSTANDING" & vbCrLf
                    mystr = mystr & "" & vbCrLf
                    mystr = mystr & FullStrip
                    sw.WriteLine(mystr)


                    For Each row2 In MyDsTable2.Rows
                        mystr = ""
                        If iCabang2 = CStr(row2("BRANCH_NAME")) Then

                            mystr = StringBuilder(7, "", row2("GL_ACCOUNT_ID"))
                            mystr = StringBuilder(75, mystr, CDec(row2("GL_OUTSTANDING")).ToString("#,#0.00", current).PadLeft(20, " "))


                            sw.WriteLine(mystr)

                        End If

                    Next
                    sw.WriteLine(FullStrip)
                    mystr = Chr(12)
                    sw.WriteLine(mystr)
                End If

            End Using
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.ToString & vbCrLf)
        End Try
    End Sub
    Function StringBuilder(ByVal myColl As Integer, ByVal str_ori As String, ByVal str As String) As String
        Dim tmp_str As String = ""
        Dim tmp_space As String = "                                                                                                                                                                                                                                                                                       "
        If str_ori = "" Then
            tmp_str = tmp_space
        Else
            tmp_str = str_ori
        End If

        StringBuilder = (tmp_str.Substring(0, myColl - 1) & str & tmp_space).Substring(0, 280)
    End Function
    Function GetAppKey(ByVal myKey As String) As String
        Try
            Dim appSettings = ConfigurationManager.AppSettings
            GetAppKey = appSettings(myKey)
        Catch e As ConfigurationErrorsException
            Console.WriteLine("Error reading app settings")
            GetAppKey = ""
        End Try
    End Function
    Sub FTPSent(ByVal FTPtransfer As String, ByVal FTPdirek As String, ByVal UserID As String, ByVal Pass As String)
        Try
            Dim request As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create(FTPtransfer), System.Net.FtpWebRequest)
            request.Proxy = Nothing
            request.Credentials = New System.Net.NetworkCredential(UserID, Pass)
            request.Method = System.Net.WebRequestMethods.Ftp.UploadFile

            Dim file() As Byte = System.IO.File.ReadAllBytes(FTPdirek)

            Dim strz As System.IO.Stream = request.GetRequestStream()
            strz.Write(file, 0, file.Length)
            strz.Close()
            strz.Dispose()


            Console.WriteLine("FTP Done...")
            'Console.ReadKey()
        Catch ex As Exception
            Console.WriteLine("FTPSent error: " & ex.ToString)
        End Try
    End Sub
    Sub CreateZipFile(ByVal MyFile As String)
        Try
            Dim fi As New FileInfo(MyFile)
            Using inFile As FileStream = fi.OpenRead()
                ' Compressing:
                ' Prevent compressing hidden and already compressed files.
                Console.WriteLine("Compressing File...")
                If (File.GetAttributes(fi.FullName) And FileAttributes.Hidden) _
                    <> FileAttributes.Hidden And fi.Extension <> ".gz" Then
                    ' Create the compressed file.
                    Using outFile As FileStream = File.Create(fi.FullName + ".gz")
                        Using Compress As GZipStream = _
                            New GZipStream(outFile, CompressionMode.Compress)

                            ' Copy the source file into the compression stream.
                            inFile.CopyTo(Compress)

                            Console.WriteLine("Compressed {0} from {1} to {2} bytes.", _
                                              fi.Name, fi.Length.ToString(), outFile.Length.ToString())

                        End Using
                    End Using
                End If
            End Using
        Catch ex As Exception
            Console.WriteLine("CreateZipFile error: " & ex.ToString)
        End Try
    End Sub
End Module
