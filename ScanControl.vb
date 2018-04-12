Imports DTI.ImageMan.Twain
Imports PdfSharp
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports System.IO
Imports System.Data.SqlClient

'/f "3" "IINV" "../../Scandocs/demo/IINV" "T8" 
'/f  "1 PORD ../../scandocs/demo/PORD PO12000017"
Public Class ScanControl
    ' Single Scan Console
    ' Version 1.0
    ' Andy Mackintosh
    ' 14/12/2012
    ' This application is intended to be used from within Priority. It accepts one parameter /f {filename with full path but no filetype}
    ' Once ran it will call the scanner and grab whatever document is loaded, then it will convert it to pdf and save it as the filename provided
    ' I have set the needed scanner contols up as application level variables in the ScanSettings class, these are read from an XML file and can
    ' therefore be changed by simply editing the file
    '
    ' Version 2.
    ' Si - 04/18
    ' Fixed it. Pretty much every single thing.

    Public Class ScanDocument

#Region "Properties"

        Private SCAN_DATE As Integer
        Public Property doc_date() As Integer
            Get
                Return SCAN_DATE
            End Get
            Set(ByVal value As Integer)
                SCAN_DATE = value
            End Set
        End Property

        Private SCAN_PROCESSED As Char
        Public Property processed() As Char
            Get
                Return SCAN_PROCESSED
            End Get
            Set(ByVal value As Char)
                SCAN_PROCESSED = value
            End Set
        End Property

        Private SCAN_BATCH_NO As Integer
        Public Property batch_no() As Integer
            Get
                Return SCAN_BATCH_NO
            End Get
            Set(ByVal value As Integer)
                SCAN_BATCH_NO = value
            End Set
        End Property

        Private _OutputFile As FileInfo
        Public ReadOnly Property OutputFile() As FileInfo
            Get
                With _OutputFile.Directory
                    If Not .Exists Then .Create()
                End With
                Return _OutputFile
            End Get
        End Property

#End Region

#Region "Constructor"

        Public Sub New(ByVal dir As String, ByVal name As String, ByVal proc As Char, ByVal bno As Integer)

            processed = proc
            batch_no = bno

            doc_date = DateDiff( _
                DateInterval.Minute, _
                #1/1/1988#, _
                Now _
            )

            Dim i As Integer = 1
            Do
                _OutputFile = New FileInfo( _
                    Path.Combine( _
                        dir, _
                        String.Format( _
                            "{0}_{1}_{2}.pdf", _
                            name, _
                            Format(Now, "yyyyMMdd"), _
                            i.ToString _
                        ) _
                    ) _
                )
                i += 1
            Loop While _OutputFile.Exists

        End Sub

#End Region

    End Class

#Region "Shared Methods"

    Public Shared Function scannall(ByVal f As ScanDocument)

        Dim imgs As New List(Of System.Drawing.Image)
        Dim img As System.Drawing.Image
        Dim D As New UserSettings

        Try
            Using tw As New DTI.ImageMan.Twain.TwainControl
                With tw
                    .SelectScanner()

                    .PixelType = D.PixType
                    .MaxPages = D.MaxPage
                    .UserInterface = UserInterfaces.None
                    .Resolution = D.Res

                    img = .ScanPage()
                    While Not (img Is Nothing)
                        imgs.Add(img)
                        img = .ScanPage()

                    End While

                End With

            End Using

            If imgs.Count = 0 Then Throw New Exception("Nothing scanned.")

            Using doc As New PdfDocument

                For Each i As System.Drawing.Image In imgs
                    'add a blank page to the pdf                                        
                    Using xgr As XGraphics = XGraphics.FromPdfPage(doc.AddPage())
                        'draw the image to the page
                        xgr.DrawImage(XImage.FromGdiPlusImage(i), 0, 0)

                    End Using

                Next

                With f
                    doc.Save(.OutputFile.FullName)
                    write_log( _
                        .OutputFile.Directory.FullName & "\", _
                        .OutputFile.Name, _
                        .doc_date, _
                        "Single", _
                        .OutputFile.Name, _
                        .batch_no, _
                        D _
                    )

                End With

            End Using

            Return True

        Catch ex As Exception
            write_error(ex.Message, 1, D)
            Return False

        End Try

    End Function

#Region "Logging"

    Public Shared Sub write_error(ByVal errmsg As String, ByVal usr As Integer, ByVal d As UserSettings)

        Using con As SqlConnection = d.ConStr
            con.Open()
            Using cmd As New SqlCommand
                With cmd
                    .Connection = con
                    .CommandText = "insert into ZEMG_ERRMSGLOG (LOGGEDBYPROCNAME,LOGGEDDATE,[MESSAGE],T$USER) values ( 'Scanner',@ddate,@message,@user)"
                    .Parameters.AddWithValue("ddate", DateDiff(DateInterval.Minute, #1/1/1988#, Now))
                    .Parameters.AddWithValue("message", errmsg)
                    .Parameters.AddWithValue("user", usr)
                    .ExecuteNonQuery()
                End With

            End Using

        End Using

    End Sub

    Public Shared Sub write_log(ByVal DOC_DIR As String, ByVal File_Name As String, ByVal S_Date As Integer, ByVal sc_ty_co As String, ByVal SDC As String, ByVal S_Batch_NO As Integer, ByVal d As UserSettings)

        Using con As SqlConnection = d.ConStr
            con.Open()
            Using cmd As New SqlCommand
                With cmd
                    .Connection = con
                    .CommandText = "insert into ZEMG_SCANNINGLOG (SCAN_DOC_DIR,SCAN_FILE_NAME,SCAN_DATE,SCAN_PROCESSED,SCAN_TYPE_CODE,SCAN_DOC_CODE,SCAN_BATCH_NO) values ( @docdir,@filename ,@ddate,'N',@TYCODE,@SDCO,@batchno)"
                    .Parameters.AddWithValue("ddate", S_Date)
                    .Parameters.AddWithValue("docdir", DOC_DIR)
                    .Parameters.AddWithValue("filename", File_Name)
                    .Parameters.AddWithValue("batchno", S_Batch_NO)
                    .Parameters.AddWithValue("TYCODE", sc_ty_co)
                    .Parameters.AddWithValue("SDCO", SDC)
                    .ExecuteNonQuery()

                End With

            End Using

        End Using

    End Sub

#End Region

#End Region

End Class
