Imports System.IO
Imports System.Data.Sql
Imports System.Data.SqlClient

Public Class UserSettings

#Region "class"

#Region "Properties"

    Public ReadOnly Property SettingsFile() As FileInfo
        Get
            Static fi As FileInfo = Nothing
            If fi Is Nothing Then
                Dim di As New DirectoryInfo( _
                    String.Format( _
                        "{0}\medatech\", _
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) _
                    ) _
                )
                With di
                    If Not .Exists Then .Create()
                    fi = New FileInfo(Path.Combine(di.FullName, "settings.xml"))

                End With
            End If
            Return fi
        End Get
    End Property

    Public Property PixType() As DTI.ImageMan.Twain.PixelTypes
        Get
            Return doc.<Settings>.<PixType>.Value
        End Get
        Set(ByVal value As DTI.ImageMan.Twain.PixelTypes)
            doc.<Settings>.<PixType>.Value = value
        End Set
    End Property

    Public Property MaxPage() As Integer
        Get
            Return doc.<Settings>.<MaxPage>.Value
        End Get
        Set(ByVal value As Integer)
            doc.<Settings>.<MaxPage>.Value = value
        End Set
    End Property

    Public Property UsrIntFc() As DTI.ImageMan.Twain.UserInterfaces
        Get
            Return doc.<Settings>.<UsrIntFc>.Value
        End Get
        Set(ByVal value As DTI.ImageMan.Twain.UserInterfaces)
            doc.<Settings>.<UsrIntFc>.Value = value
        End Set
    End Property

    Public Property Res() As Integer
        Get
            Return doc.<Settings>.<Res>.Value
        End Get
        Set(ByVal value As Integer)
            doc.<Settings>.<Res>.Value = value
        End Set
    End Property

    Public Property DBServ() As String
        Get
            Return doc.<Settings>.<DBServ>.Value
        End Get
        Set(ByVal value As String)
            doc.<Settings>.<DBServ>.Value = value
        End Set
    End Property

    Public Property DBName() As String
        Get
            Return doc.<Settings>.<DBName>.Value
        End Get
        Set(ByVal value As String)
            doc.<Settings>.<DBName>.Value = value
        End Set
    End Property

    Public Property DBUname() As String
        Get
            Return doc.<Settings>.<DBUname>.Value
        End Get
        Set(ByVal value As String)
            doc.<Settings>.<DBUname>.Value = value
        End Set
    End Property

    Public Property DBPassword() As String
        Get
            Return doc.<Settings>.<DBPass>.Value
        End Get
        Set(ByVal value As String)
            doc.<Settings>.<DBPass>.Value = value
        End Set
    End Property

    Public ReadOnly Property ConStr() As SqlConnection
        Get
            Return New SqlConnection( _
                String.Format( _
                    "Data Source={0} ;Uid={1};Pwd={2};Initial Catalog={3}", _
                    DBServ, _
                    DBUname, _
                    DBPassword, _
                    DBName _
                ) _
            )
        End Get
    End Property

#End Region

#Region "Constructor"

    Private doc As XDocument
    Public Sub New(Optional ByVal Create As Boolean = False)

        InitializeComponent()

        If Not SettingsFile.Exists Then
            If Not Create Then Throw New Exception("Missing settings file.")

            Dim xmldoc As New Xml.XmlDocument
            Dim headelement As Xml.XmlElement = xmldoc.CreateElement("Settings")

            xmldoc.AppendChild(headelement)

            Dim e1 As Xml.XmlElement = xmldoc.CreateElement("PixType")
            e1.InnerText = 0
            headelement.AppendChild(e1)

            Dim el2 As Xml.XmlElement = xmldoc.CreateElement("MaxPage")
            el2.InnerText = -1
            headelement.AppendChild(el2)

            Dim el3 As Xml.XmlElement = xmldoc.CreateElement("UsrIntFc")
            el3.InnerText = 1
            headelement.AppendChild(el3)

            Dim el4 As Xml.XmlElement = xmldoc.CreateElement("Res")
            el4.InnerText = 72
            headelement.AppendChild(el4)

            Dim el5 As Xml.XmlElement = xmldoc.CreateElement("DBServ")
            el5.InnerText = ""
            headelement.AppendChild(el5)

            Dim el6 As Xml.XmlElement = xmldoc.CreateElement("DBUname")
            el6.InnerText = ""
            headelement.AppendChild(el6)

            Dim el7 As Xml.XmlElement = xmldoc.CreateElement("DBPass")
            el7.InnerText = ""
            headelement.AppendChild(el7)

            Dim el8 As Xml.XmlElement = xmldoc.CreateElement("DBName")
            el8.InnerText = ""
            headelement.AppendChild(el8)

            xmldoc.Save(SettingsFile.FullName)

            doc = XDocument.Load(SettingsFile.FullName)

        Else
            doc = XDocument.Load(SettingsFile.FullName)

            selectPixType.SelectedIndex = PixType
            selectMaxPage.Value = MaxPage
            selectUsrIntFc.SelectedIndex = UsrIntFc
            selectRes.Value = Res
            selectDBServ.Text = DBServ
            selectDBName.Text = DBName

            txtDBUname.Text = DBUname
            txtDBPassword.Text = DBPassword

        End If

        AddHandler selectRes.ValueChanged, AddressOf selectRes_ValueChanged
        AddHandler selectPixType.SelectedIndexChanged, AddressOf selectPixType_SelectedIndexChanged
        AddHandler selectMaxPage.ValueChanged, AddressOf selectMaxPage_ValueChanged
        AddHandler selectUsrIntFc.SelectedIndexChanged, AddressOf selectUsrIntFc_SelectedIndexChanged

    End Sub

#End Region

#End Region

#Region "Form"

    Dim filling As Boolean = False
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim instance As SqlDataSourceEnumerator = _
            SqlDataSourceEnumerator.Instance
        Dim table As System.Data.DataTable = instance.GetDataSources()
        Dim listOfServers As New List(Of String)()
        For Each RowOfData As DataRow In table.Rows
            'get the server name 
            Dim serverName As String = RowOfData("ServerName").ToString()
            'get the instance name 
            Dim instanceName As String = RowOfData("InstanceName").ToString()

            'check if the instance name is empty 
            If Not instanceName.Equals(String.Empty) Then
                'append the instance name to the server name 
                serverName += String.Format("\{0}", instanceName)
            End If
            'add the server to our list 

            listOfServers.Add(serverName)


        Next
        filling = True
        selectDBServ.DataSource = listOfServers
        filling = False


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles selectDBServ.SelectedIndexChanged
        If filling = True Then Exit Sub
        Dim listDataBases As List(Of String) = New List(Of String)
        Dim connectString As String        
        Dim server As String
        ' Check if user was selected a server to connect
        selectDBName.DataSource = Nothing
        If selectDBServ.Text = "" Then
            MsgBox("Must select a server")
            Exit Sub
        Else
            server = selectDBServ.Text
        End If
        If RadioButton1.Checked = True Then
            connectString = "Data Source=" & server & " ;Integrated Security=True;Initial Catalog=master"
        Else
            If txtDBUname.Text = "" Or txtDBPassword.Text = "" Then
                MsgBox("Must Provide username and password")
                Exit Sub
                'Uid=myUsername;Pwd=myPassword;

            End If
            connectString = "Data Source=" & server & " ;Uid=" & txtDBUname.Text & ";Pwd=" & txtDBPassword.Text & ";Initial Catalog=master"
        End If



        Using con As New SqlConnection(connectString)
            ' Open connection
            Try
                con.Open()
                'Get databases names in server in a datareader                
                Dim com As SqlCommand = New SqlCommand("select name from sys.databases;", con)
                Dim dr As SqlDataReader = com.ExecuteReader()
                While (dr.Read())
                    listDataBases.Add(dr(0).ToString())
                End While
                'Set databases list as combobox’s datasource 
                selectDBName.DataSource = listDataBases

            Catch ex As Exception
                MsgBox("These settings generated an error, please check and try again")
            End Try

        End Using


    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            selectDBName.DataSource = Nothing
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If selectDBServ.Text = "" Then
            MsgBox("No server chosen please check and try again")
            Exit Sub
        End If
        If RadioButton2.Checked = True Then
            If txtDBUname.Text = "" Then
                MsgBox("No username selected")
                Exit Sub
            End If
            If txtDBPassword.Text = "" Then
                MsgBox("No password selected")
                Exit Sub
            End If

        End If
        If selectDBName.Text = "" Then
            MsgBox("No database selected please check and try again")
        End If

        Dim connectstring As String
        If RadioButton1.Checked = True Then
            connectstring = "Data Source=" & selectDBServ.Text & " ;Integrated Security=True;Initial Catalog=" & selectDBName.Text
        Else
            connectstring = "Data Source=" & selectDBServ.Text & " ;Uid=" & txtDBUname.Text & ";Pwd=" & txtDBPassword.Text & ";Initial Catalog=" & selectDBName.Text
        End If

        Using con As New SqlConnection(connectstring)
            ' Open connection
            Try
                con.Open()
                If con.State = ConnectionState.Open Then
                    MsgBox("Test Succesful")
                    CheckBox1.Checked = True
                End If
            Catch ex As Exception
                MsgBox("These settings generated an error, please check and try again")
            End Try

        End Using



    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        If selectDBServ.Text = "" Then Exit Sub
        If selectDBName.Text = "" Then Exit Sub
        If selectUsrIntFc.Text = "" Then Exit Sub
        If selectPixType.Text = "" Then Exit Sub
        If txtDBUname.Text = "" Then Exit Sub
        If txtDBPassword.Text = "" Then Exit Sub

        If CheckBox1.Checked = False Then
            If Not MsgBox("Database settings havent been verified. Save anyway?", MsgBoxStyle.OkCancel + MsgBoxStyle.DefaultButton2, "Proceed?") = MsgBoxResult.Ok Then
                Exit Sub
            End If
        End If

        With Me
            .DBServ = .selectDBServ.Text
            .DBName = .selectDBName.Text
            .DBUname = txtDBUname.Text
            .DBPassword = txtDBPassword.Text

        End With

        doc.Save(SettingsFile.FullName)
        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub


#End Region

#Region "Control Handlers"

    Private Sub selectPixType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Select Case selectPixType.Text
            Case "Black and White"
                PixType = 0
            Case "Greyscale"
                PixType = 1
            Case "RGB Colour"
                PixType = 2
            Case "Palette Colour"
                PixType = 3

        End Select

    End Sub

    Private Sub selectMaxPage_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MaxPage = selectMaxPage.Value

    End Sub

    Private Sub selectUsrIntFc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Select Case selectUsrIntFc.Text
            Case "Show Interface"
                UsrIntFc = 1

            Case "Hide Interface"
                UsrIntFc = 0

            Case "Modal Interface"
                UsrIntFc = 2

        End Select

    End Sub

    Private Sub selectRes_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Res = selectRes.Value

    End Sub

#End Region

End Class