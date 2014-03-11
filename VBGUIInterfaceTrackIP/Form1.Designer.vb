Imports System.Threading
Imports System.Net
Imports System.Text
Imports System.IO


<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TheForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Private trd As Thread
    Private globalTemp As String
    Private docCompleteHandler = New WebBrowserDocumentCompletedEventHandler(AddressOf TraceRouteThread)



    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Browser = New System.Windows.Forms.WebBrowser
        Me.EntryBox = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TraceButton = New System.Windows.Forms.Button
        Me.HomeButton = New System.Windows.Forms.Button
        Me.Status = New System.Windows.Forms.StatusStrip
        Me.Information = New System.Windows.Forms.WebBrowser
        Me.SuspendLayout()
        '
        'Browser
        '
        Me.Browser.Location = New System.Drawing.Point(0, 26)
        Me.Browser.MinimumSize = New System.Drawing.Size(20, 20)
        Me.Browser.Name = "Browser"
        Me.Browser.ScriptErrorsSuppressed = True
        Me.Browser.Size = New System.Drawing.Size(781, 352)
        Me.Browser.TabIndex = 0
        Me.Browser.Url = New System.Uri("http://www.reddit.com", System.UriKind.Absolute)
        '
        'EntryBox
        '
        Me.EntryBox.Location = New System.Drawing.Point(76, 3)
        Me.EntryBox.Name = "EntryBox"
        Me.EntryBox.Size = New System.Drawing.Size(241, 20)
        Me.EntryBox.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(2, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "IP Address"
        '
        'TraceButton
        '
        Me.TraceButton.Location = New System.Drawing.Point(323, 1)
        Me.TraceButton.Name = "TraceButton"
        Me.TraceButton.Size = New System.Drawing.Size(75, 23)
        Me.TraceButton.TabIndex = 3
        Me.TraceButton.Text = "Trace IP"
        Me.TraceButton.UseVisualStyleBackColor = True

        '
        'HomeButton
        '
        Me.HomeButton.Location = New System.Drawing.Point(400, 1)
        Me.HomeButton.Name = "HomeButton"
        Me.HomeButton.Size = New System.Drawing.Size(75, 23)
        Me.HomeButton.TabIndex = 4
        Me.HomeButton.Text = "Home"
        Me.HomeButton.UseVisualStyleBackColor = True

        '
        'Status
        '
        Me.Status.Location = New System.Drawing.Point(0, 603)
        Me.Status.Name = "Status"
        Me.Status.Size = New System.Drawing.Size(781, 22)
        Me.Status.TabIndex = 5
        Me.Status.Text = "StatusStrip1"
        '
        'Information
        '
        Me.Information.Location = New System.Drawing.Point(0, 384)
        Me.Information.Name = "Information"
        Me.Information.Size = New System.Drawing.Size(781, 216)
        Me.Information.ScriptErrorsSuppressed = True
        Me.Information.TabIndex = 6
        'Me.Information.Text = ""
        '
        'TheForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(781, 625)
        Me.Controls.Add(Me.Information)
        Me.Controls.Add(Me.Status)
        Me.Controls.Add(Me.TraceButton)
        Me.Controls.Add(Me.HomeButton)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.EntryBox)
        Me.Controls.Add(Me.Browser)
        Me.Name = "TheForm"
        Me.Text = "I'll create a GUI interface using Visual Basic. See if I can track an IP address."
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Browser As System.Windows.Forms.WebBrowser
    Friend WithEvents EntryBox As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TraceButton As System.Windows.Forms.Button
    Friend WithEvents HomeButton As System.Windows.Forms.Button
    Friend WithEvents Status As System.Windows.Forms.StatusStrip
    Friend WithEvents Information As System.Windows.Forms.WebBrowser

    Private Sub TheForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TheForm_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Browser.Width = MyBase.Size.Width - 10
        Browser.Height = Math.Floor(MyBase.Size.Height * 0.65)
        Information.Width = MyBase.Size.Width - 10
        Information.Location = New Point(Information.Location.X, Browser.Location.Y + Browser.Size.Height + 3)
        Information.Height = MyBase.Size.Height - ((2 * Browser.Location.Y) + Browser.Size.Height + Status.Size.Height + 10)
    End Sub
    
    Private Sub TraceButton_Click(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles TraceButton.MouseClick
        Dim strHTML As String
        strHTML = GetPageHTML("http://samspade.org/whois/" & EntryBox.Text)
        Dim intLocation, intDivClose As Integer
        Dim theDiv As String
        theDiv = "<div class=""set"">"
        intLocation = InStr(strHTML, theDiv) + theDiv.Length
        intDivClose = InStr(intLocation, strHTML, "</div>")
        strHTML = Mid(strHTML, intLocation, intDivClose - intLocation)

        ' Add an event handler that prints the document after it loads.
        AddHandler Information.DocumentCompleted, docCompleteHandler

        globalTemp = strHTML
        Information.DocumentText = strHTML


        Console.WriteLine("http://samspade.org/whois/" & EntryBox.Text)

        ' Create a request using a URL that can receive a post. 
        Dim request As WebRequest = WebRequest.Create("http://www.geobytes.com/IpLocator.htm?GetLocation")
        ' Set the Method property of the request to POST.
        request.Method = "POST"

        ' Create POST data and convert it to a byte array.
        Dim postData As String = "ipaddress=" & EntryBox.Text & "&cid=0&c=0&Template=iplocator.htm"
        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)
        ' Set the ContentType property of the WebRequest.
        request.ContentType = "application/x-www-form-urlencoded"
        ' Set the ContentLength property of the WebRequest.
        request.ContentLength = byteArray.Length
        ' Get the request stream.
        Dim dataStream As Stream = request.GetRequestStream()
        ' Write the data to the request stream.
        dataStream.Write(byteArray, 0, byteArray.Length)
        ' Close the Stream object.
        dataStream.Close()
        ' Get the response.
        Dim response As WebResponse = request.GetResponse()
        ' Display the status.
        Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
        ' Get the stream containing content returned by the server.
        dataStream = response.GetResponseStream()
        ' Open the stream using a StreamReader for easy access.
        Dim reader As New StreamReader(dataStream)
        ' Read the content.
        Dim responseFromServer As String = reader.ReadToEnd()
        ' Display the content.
        'vvvvvv change this vvvvvv
        Browser.ScriptErrorsSuppressed = True
        ' Handle DocumentCompleted to gain access to the Document object.
        AddHandler Browser.DocumentCompleted, AddressOf browser_DocumentCompleted

        Dim formStart, formEnd As String
        formStart = "<form method=""POST"" action=""http://localhost/IpLocator.htm?GetLocation"">" & vbNewLine & _
                    " <!-- If you would like direct access to our map database,"
        Dim startLocation, endLocation As Integer
        startLocation = InStr(responseFromServer, formStart) + _
                        Len("<form method=""POST"" action=""http://localhost/IpLocator.htm?GetLocation"">" & vbNewLine & _
                            " <!-- If you would like direct access to our map database, then please visit www.geobytes.com" & vbNewLine & _
                            "  ---- We have developer license from $49 which works out a lot cheaper then trying to rip" & vbNewLine & _
                            "  ---- all of our data through this page :) -->") + 1
        formEnd = "<tr></form>"
        endLocation = InStr(Mid(responseFromServer, startLocation), formEnd) - 1

        responseFromServer = Mid(responseFromServer, startLocation, endLocation) & "</table>"

        Browser.DocumentText = responseFromServer
        Console.WriteLine(responseFromServer)

        ' Clean up the streams.
        reader.Close()
        dataStream.Close()
        response.Close()



        'Browser.Navigate(New Uri("http://samspade.org/whois/" & EntryBox.Text))
    End Sub

    Private Sub HomeButton_Click(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles HomeButton.MouseClick
        Console.WriteLine("http://www.reddit.com")
        Browser.Navigate(New Uri("http://www.reddit.com"))
    End Sub

    Public Function GetPageHTML(ByVal URL As String) As String
        ' Retrieves the HTML from the specified URL
        Dim objWC As New System.Net.WebClient()
        Return New System.Text.UTF8Encoding().GetString(objWC.DownloadData(URL))
    End Function

    Private Sub TraceRoute()
        Dim strHTML As String
        Dim intLocation As Integer
        Dim intDivClose As Integer

        strHTML = GetPageHTML("http://www.tracert.org/cgi-bin/traceroute/tracert.pl?t=" & Thread.CurrentThread.Name)
        intLocation = InStr(strHTML, "<PRE>") + "<PRE>".Length
        intDivClose = InStr(intLocation, strHTML, "</PRE>")
        strHTML = Mid(strHTML, intLocation, intDivClose - intLocation)
        Information.DocumentText = globalTemp & vbNewLine & vbNewLine & "<pre>" & strHTML & "</pre>" & vbNewLine
    End Sub

    Private Sub TraceRouteThread(sender As System.Object, e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs)
        trd = New Thread(AddressOf TraceRoute)
        trd.IsBackground = True
        trd.Name = EntryBox.Text
        trd.Start()
        RemoveHandler Information.DocumentCompleted, docCompleteHandler
    End Sub


    
    Private Sub browser_DocumentCompleted(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs)
        AddHandler CType(sender, WebBrowser).Document.Window.Error, AddressOf Window_Error
    End Sub

    Private Sub Window_Error(ByVal sender As Object, ByVal e As HtmlElementErrorEventArgs)
        ' Ignore the error and suppress the error dialog box. 
        e.Handled = True
    End Sub



End Class
