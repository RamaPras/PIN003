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

            'Create Text File
            Dim myFileName As String = CreateFileName(args, GetAppKey("FILENAME"))
            CreateFile(MyDataSet, GetAppKey("DIRECTORY") & myFileName)

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
    Sub CreateFile(ByVal MyDataSet As DataSet, ByVal MyFileName As String)
        Try
            Dim MyDsTable = MyDataSet.Tables(0)
            Dim columnCount = MyDsTable.Columns.Count
            Console.WriteLine("Create Textfile: " & MyFileName & vbCrLf)
            'Console.WriteLine("Column Count: " & columnCount)

            Dim row As DataRow
            Console.WriteLine()
            Dim mystr As String = ""
            Dim mystr2 As String = ""
            Dim counter As Integer
            Using sw As StreamWriter = New StreamWriter(MyFileName)
                counter = 1

                Dim strDlm As String = GetAppKey("DELIMITER")
                Dim iCabang As String = ""
                Dim iCabang2 As String = ""
                Dim Strip As String = ""
                Dim FullStrip As String = ""
                Dim iPage As Integer = 1
                Dim iNomor As Integer = 0
                Dim Pemisah As Integer = 0
                Dim Pemisah2 As Integer = 0
                Dim IDX As Integer = 0
                Dim Cont As Integer = 0
                Dim HeadPage As Boolean = True
                Dim AccType As Boolean = True
                Dim First As Boolean = True
                Dim First2 As Boolean = True
                Dim First3 As Boolean = True
                Dim First4 As Boolean = True
                Dim First5 As Boolean = True
                Dim First6 As Boolean = True
                Dim TOT_MAKSIMUM_KREDIT As Decimal
                Dim TOT_IZIN_TARIK As Decimal
                Dim TOT_SALDO_AKHIR_PLUS As Decimal
                Dim TOT_POKOK_TUNGGAKAN As Decimal
                Dim TOT_BUNGA As Decimal
                Dim TOT_SALDO_DENDA As Decimal
                Dim TOT_DISPONIBLE As Decimal
                Dim TOT_SALDO_AKHIR_MIN As Decimal
                Dim current As CultureInfo = New System.Globalization.CultureInfo("id-ID")
                'Console.WriteLine("The current UI culture is {0}", current.Name)
                FullStrip = "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                For Each row In MyDsTable.Rows
                    If First5 = True Then
                        iCabang2 = CStr(row("NAMA_CABANG"))
                        iCabang = iCabang2
                        First5 = False
                    Else
                        iCabang = CStr(row("NAMA_CABANG"))
                    End If
                    If iCabang <> "" Then
ATAS:
                        Cont = CStr(row("CONT"))
                        Pemisah = CInt(row("IDX"))
                        If (Pemisah = 1 Or Pemisah = 4) Then
                            Strip = "                                                                                                    --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                        ElseIf (Pemisah = 2 Or Pemisah = 3 Or Pemisah = 5 Or Pemisah = 6) Then
                            Strip = "                                                                                                    ---------------------------------------------------------------------------------------------------------------------------------------------------------------------"
                        End If
                        If HeadPage = True Then
                            If iPage <> 1 Then
                                mystr = Chr(12) & "PT BANK NEGARA INDONESIA (Persero) Tbk                                                                                                                                                                                                           PIN003H2"
                                sw.WriteLine(mystr)
                            Else
                                mystr = "PT BANK NEGARA INDONESIA (Persero) Tbk                                                                                                                                                                                                           PIN003H2"
                                sw.WriteLine(mystr)
                            End If
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
                            mystr = StringBuilder(120, "", "SALDOLIST PINJAMAN")
                            sw.WriteLine(mystr)
                            sw.WriteLine("")
                            mystr = StringBuilder(119, "", "Tanggal : " & row("TANGGAL"))
                            sw.WriteLine(mystr)
                            sw.WriteLine("")
                            sw.WriteLine("")
                            mystr = FullStrip & vbCrLf
                            mystr = mystr & "NO.    NO CIF                NO.            NAMA DEBITUR      GLCC    SEG    JENIS   PML BI KOL  JW        MAKSIMUM              IZIN TARIK            SALDO AKHIR                            TUNGGAKAN                         SALDO DENDA              DISPONIBLE         UTIL REK" & vbCrLf
                            mystr = mystr & "                          REKENING                                          PGUNAAN                         KREDIT                                                                  POKOK                  BUNGA                                                                    " & vbCrLf
                            mystr = mystr & FullStrip
                            sw.WriteLine(mystr)
                            IDX = IDX + 13
                            HeadPage = False
                        End If

                        If First4 = False Then
                            If iCabang <> iCabang2 Then
                                If IDX <> 0 Then
                                    iPage = iPage + 1
                                End If
                                sw.WriteLine("")
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

                        TOT_MAKSIMUM_KREDIT = CDec(row("TOT_MAKSIMUM_KREDIT"))
                        TOT_IZIN_TARIK = CDec(row("TOT_IZIN_TARIK"))
                        TOT_SALDO_AKHIR_PLUS = CDec(row("TOT_SALDO_AKHIR_PLUS"))
                        TOT_POKOK_TUNGGAKAN = CDec(row("TOT_POKOK_TUNGGAKAN"))
                        TOT_BUNGA = CDec(row("TOT_BUNGA"))
                        TOT_SALDO_DENDA = CDec(row("TOT_SALDO_DENDA"))
                        TOT_DISPONIBLE = CDec(row("TOT_DISPONIBLE"))
                        TOT_SALDO_AKHIR_MIN = CDec(row("TOT_SALDO_AKHIR_MIN"))

                        If (Pemisah = 1 Or Pemisah = 4) Then
                            mystr2 = ""
                            mystr2 = StringBuilder(1, "", CStr(iNomor + 1) & ".")
                            mystr2 = StringBuilder(8, mystr2, Trim(CStr(row("NO_CIF"))))
                            mystr2 = StringBuilder(23, mystr2, Trim(CStr(row("NO_REKENING"))))
                            mystr2 = StringBuilder(41, mystr2, Trim(CStr(row("NAMA_DEBITUR"))))
                            mystr2 = StringBuilder(63, mystr2, Trim(CStr(row("GLCC"))))
                            mystr2 = StringBuilder(71, mystr2, Trim(CStr(row("SEG"))))
                            mystr2 = StringBuilder(83, mystr2, Trim(CStr(row("JENIS_PGUNAAN"))))
                            mystr2 = StringBuilder(86, mystr2, Trim(CStr(row("PML_BI"))))
                            mystr2 = StringBuilder(94, mystr2, Trim(CStr(row("KOL"))))
                            mystr2 = StringBuilder(98, mystr2, Trim(CStr(row("JW"))))
                            mystr2 = StringBuilder(101, mystr2, Trim(Replace(Replace(Replace(CDec(row("MAKSIMUM_KREDIT")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(23, " "))
                            mystr2 = StringBuilder(125, mystr2, Trim(Replace(Replace(Replace(CDec(row("IZIN_TARIK")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr2 = StringBuilder(148, mystr2, Trim(Replace(Replace(Replace(CDec(row("SALDO_AKHIR")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr2 = StringBuilder(171, mystr2, Trim(Replace(Replace(Replace(CDec(row("POKOK_TUNGGAKAN")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(25, " "))
                            mystr2 = StringBuilder(197, mystr2, Trim(Replace(Replace(Replace(CDec(row("BUNGA")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(21, " "))
                            mystr2 = StringBuilder(219, mystr2, Trim(Replace(Replace(Replace(CDec(row("SALDO_DENDA")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr2 = StringBuilder(242, mystr2, Trim(Replace(Replace(Replace(CDec(row("DISPONIBLE")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(24, " "))
                            mystr2 = StringBuilder(267, mystr2, Trim(Replace(Replace(Replace(CDec(row("UTIL_REK")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(10, " "))

                            If AccType = True Then
                                mystr = "        " & row("ACCT_TYPE") & " - " & row("NAMA_PROD")
                                sw.WriteLine(mystr)
                                IDX = IDX + 1
                                AccType = False
                            End If

                            mystr = StringBuilder(1, "", CStr(iNomor + 1) & ".")
                            mystr = StringBuilder(8, mystr, Trim(CStr(row("NO_CIF"))))
                            mystr = StringBuilder(23, mystr, Trim(CStr(row("NO_REKENING"))))
                            mystr = StringBuilder(41, mystr, Trim(CStr(row("NAMA_DEBITUR"))))
                            mystr = StringBuilder(63, mystr, Trim(CStr(row("GLCC"))))
                            mystr = StringBuilder(71, mystr, Trim(CStr(row("SEG"))))
                            mystr = StringBuilder(83, mystr, Trim(CStr(row("JENIS_PGUNAAN"))))
                            mystr = StringBuilder(86, mystr, Trim(CStr(row("PML_BI"))))
                            mystr = StringBuilder(94, mystr, Trim(CStr(row("KOL"))))
                            mystr = StringBuilder(98, mystr, Trim(CStr(row("JW"))))
                            mystr = StringBuilder(101, mystr, Trim(Replace(Replace(Replace(CDec(row("MAKSIMUM_KREDIT")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(23, " "))
                            mystr = StringBuilder(125, mystr, Trim(Replace(Replace(Replace(CDec(row("IZIN_TARIK")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr = StringBuilder(148, mystr, Trim(Replace(Replace(Replace(CDec(row("SALDO_AKHIR")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr = StringBuilder(171, mystr, Trim(Replace(Replace(Replace(CDec(row("POKOK_TUNGGAKAN")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(25, " "))
                            mystr = StringBuilder(197, mystr, Trim(Replace(Replace(Replace(CDec(row("BUNGA")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(21, " "))
                            mystr = StringBuilder(219, mystr, Trim(Replace(Replace(Replace(CDec(row("SALDO_DENDA")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr = StringBuilder(242, mystr, Trim(Replace(Replace(Replace(CDec(row("DISPONIBLE")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(24, " "))
                            mystr = StringBuilder(267, mystr, Trim(Replace(Replace(Replace(CDec(row("UTIL_REK")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(10, " "))
                            sw.WriteLine(mystr)
                            IDX = IDX + 1
                            iNomor = iNomor + 1
                            First2 = True
                        End If

                        If (Pemisah = 2 Or Pemisah = 3 Or Pemisah = 5 Or Pemisah = 6) Then
                            Pemisah2 = Pemisah
                            mystr2 = ""
                            If (Pemisah = 2 Or Pemisah = 5) Then
                                mystr2 = StringBuilder(24, "", row("ACCT_TYPE") & " - " & Trim(Replace(CStr(row("NAMA_PROD")), "  ", " ")))
                            ElseIf (Pemisah = 3 Or Pemisah = 6) Then
                                mystr2 = StringBuilder(23, "", row("GLCC") & " - " & row("NAMA_PROD"))
                            End If
                            mystr2 = StringBuilder(101, mystr2, Trim(Replace(Replace(Replace(CDec(row("MAKSIMUM_KREDIT")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(23, " "))
                            mystr2 = StringBuilder(125, mystr2, Trim(Replace(Replace(Replace(CDec(row("IZIN_TARIK")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr2 = StringBuilder(148, mystr2, Trim(Replace(Replace(Replace(CDec(row("SALDO_AKHIR")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr2 = StringBuilder(171, mystr2, Trim(Replace(Replace(Replace(CDec(row("POKOK_TUNGGAKAN")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(25, " "))
                            mystr2 = StringBuilder(197, mystr2, Trim(Replace(Replace(Replace(CDec(row("BUNGA")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(21, " "))
                            mystr2 = StringBuilder(219, mystr2, Trim(Replace(Replace(Replace(CDec(row("SALDO_DENDA")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr2 = StringBuilder(242, mystr2, Trim(Replace(Replace(Replace(CDec(row("DISPONIBLE")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(24, " "))
                            iNomor = iNomor + 1
                            If (Pemisah = 2 Or Pemisah = 5) Then
                                If First2 = True Then
                                    sw.WriteLine(FullStrip)
                                    IDX = IDX + 1
                                    sw.WriteLine("                      REKAPITULASI PER JENIS KREDIT:")
                                    IDX = IDX + 1
                                    First2 = False
                                    First3 = True
                                End If
                                mystr = StringBuilder(24, "", row("ACCT_TYPE") & " - " & Trim(Replace(CStr(row("NAMA_PROD")), "  ", " ")))
                            ElseIf (Pemisah = 3 Or Pemisah = 6) Then
                                If First3 = True Then
                                    sw.WriteLine(FullStrip)
                                    IDX = IDX + 1
                                    sw.WriteLine("                      REKAPITULASI PER GL CLASSIFICATION CODE:")
                                    IDX = IDX + 1
                                    First3 = False
                                End If
                                mystr = StringBuilder(23, "", row("GLCC") & " - " & row("NAMA_PROD"))
                            End If

                            mystr = StringBuilder(101, mystr, Trim(Replace(Replace(Replace(CDec(row("MAKSIMUM_KREDIT")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(23, " "))
                            mystr = StringBuilder(125, mystr, Trim(Replace(Replace(Replace(CDec(row("IZIN_TARIK")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr = StringBuilder(148, mystr, Trim(Replace(Replace(Replace(CDec(row("SALDO_AKHIR")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr = StringBuilder(171, mystr, Trim(Replace(Replace(Replace(CDec(row("POKOK_TUNGGAKAN")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(25, " "))
                            mystr = StringBuilder(197, mystr, Trim(Replace(Replace(Replace(CDec(row("BUNGA")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(21, " "))
                            mystr = StringBuilder(219, mystr, Trim(Replace(Replace(Replace(CDec(row("SALDO_DENDA")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr = StringBuilder(242, mystr, Trim(Replace(Replace(Replace(CDec(row("DISPONIBLE")).ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(24, " "))
                            sw.WriteLine(mystr)
                            IDX = IDX + 1
                            'DEBUG
                            iNomor = iNomor
                            'DEBUG
                        End If

                        If Cont = iNomor Then
                            sw.WriteLine(Strip)
                            IDX = IDX + 1

                            mystr = StringBuilder(23, "", "Total")
                            mystr = StringBuilder(101, mystr, Trim(Replace(Replace(Replace(TOT_MAKSIMUM_KREDIT.ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(23, " "))
                            mystr = StringBuilder(125, mystr, Trim(Replace(Replace(Replace(TOT_IZIN_TARIK.ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr = StringBuilder(148, mystr, Trim(Replace(Replace(Replace(TOT_SALDO_AKHIR_PLUS.ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr = StringBuilder(171, mystr, Trim(Replace(Replace(Replace(TOT_POKOK_TUNGGAKAN.ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(25, " "))
                            mystr = StringBuilder(197, mystr, Trim(Replace(Replace(Replace(TOT_BUNGA.ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(21, " "))
                            mystr = StringBuilder(219, mystr, Trim(Replace(Replace(Replace(TOT_SALDO_DENDA.ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                            mystr = StringBuilder(242, mystr, Trim(Replace(Replace(Replace(TOT_DISPONIBLE.ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(24, " "))
                            sw.WriteLine(mystr)
                            IDX = IDX + 1

                            If (Pemisah <> 3 And Pemisah <> 6) Then
                                mystr = StringBuilder(148, "", Trim(Replace(Replace(Replace(TOT_SALDO_AKHIR_MIN.ToString("C2", current), "Rp", ""), "(", "-"), ")", "")).PadLeft(22, " "))
                                sw.WriteLine(mystr)
                                IDX = IDX + 1
                            End If

                            sw.WriteLine(Strip)
                            IDX = IDX + 1

                            sw.WriteLine("")
                            IDX = IDX + 1

                            iNomor = 0
                            AccType = True
                            First6 = False
                        End If

                        If IDX > 81 Then
                            IDX = 0
                            If First = True Then
                                If First6 = True Then
                                    sw.WriteLine("")
                                End If
                                sw.WriteLine("")
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
