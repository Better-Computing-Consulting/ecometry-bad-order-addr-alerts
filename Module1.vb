Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail
Module Module1
    'Dim connString1 As String = "Data Source=ecom-db1;Initial Catalog=ECOMLIVE;Integrated Security=TRUE"
    Dim connString1 As String = "Data Source=ecom-db1;Initial Catalog=ECOMLIVE;UID=sss;PWD=sssss"
    Dim report As String = ""
   Sub Main()
      Dim afile As String = ""
      If My.Application.CommandLineArgs.Count > 0 Then
         afile = My.Application.CommandLineArgs(0)
         ProcessOneOrdersFile(afile)
      Else
         Console.WriteLine("The program needs the path of a CLEANORD file as argument.")
      End If
      Dim anEmail As New MailMessage
      With anEmail
            .From = New MailAddress("ECOM-WEB@ecommerce.com")
            .Subject = "Bad Address Orders"
            .To.Add("federico@ecommerce.com")
            If report.Length > 10 Then
                .To.Add("frank@ecommerce.com")
                .To.Add("sean@ecommerce.com")
                .To.Add("lauren@ecommerce.com")
                .Priority = MailPriority.High
            .Body = report
         End If
      End With
      Dim aSMTPClient As New SmtpClient("USA2")
      aSMTPClient.Send(anEmail)
   End Sub
   Sub ProcessOneOrdersFile(ByVal aFile As String)
      Dim newOrder As Boolean = True
      Dim shipToAddress As Boolean = False
      Dim newFile As Boolean = True
      Try
         Using sr As StreamReader = New StreamReader(aFile)
            Dim aline As String = ""
            Dim anAddrs As New OrderAddress
            Dim dbAddress As New OrderAddress
            Do
               aline = sr.ReadLine()
               If Not String.IsNullOrEmpty(aline) Then
                  If aline.StartsWith("10") Then
                     If Not newFile Then
                        dbAddress = GetAddressFromDB(anAddrs.OrderNo)
                        If Not String.IsNullOrEmpty(dbAddress.Street) Then
                           If Not inRejectFile(anAddrs.OrderNo, aFile) Then
                              If String.Compare(anAddrs.Street.Replace(".", ""), dbAddress.Street.Replace(".", ""), True) <> 0 Then
                                 Console.WriteLine("Address Problem: ")
                                 Console.WriteLine("From File: " & anAddrs.ToString)
                                 Console.WriteLine("From db:   " & dbAddress.ToString)
                                 Console.WriteLine()
                                 report &= "Address Problem: " & vbCrLf
                                 report &= "From File: " & anAddrs.ToString & vbCrLf
                                 report &= "From db:   " & dbAddress.ToString & vbCrLf & vbCrLf
                              ElseIf String.Compare(anAddrs.Zip, dbAddress.Zip, False) <> 0 Then
                                 Console.WriteLine("Address Problem: ")
                                 Console.WriteLine("From File: " & anAddrs.ToString)
                                 Console.WriteLine("From db:   " & dbAddress.ToString)
                                 Console.WriteLine()
                                 report &= "Address Problem: " & vbCrLf
                                 report &= "From File: " & anAddrs.ToString & vbCrLf
                                 report &= "From db:   " & dbAddress.ToString & vbCrLf & vbCrLf
                              End If
                           End If
                        End If
                     End If
                     newFile = False
                     newOrder = True
                     shipToAddress = False
                     anAddrs = ParseCustAddress(aline)
                  ElseIf aline.StartsWith("30") Then
                     shipToAddress = True
                     anAddrs = ParseShipToAddress(aline)
                  End If
               End If
            Loop Until aline Is Nothing
            dbAddress = GetAddressFromDB(anAddrs.OrderNo)
            If Not String.IsNullOrEmpty(dbAddress.Street) Then
               If Not inRejectFile(anAddrs.OrderNo, aFile) Then
                  If String.Compare(anAddrs.Street.Replace(".", ""), dbAddress.Street.Replace(".", ""), True) <> 0 Then
                     Console.WriteLine("Address Problem: ")
                     Console.WriteLine("From File: " & anAddrs.ToString)
                     Console.WriteLine("From db:   " & dbAddress.ToString)
                     Console.WriteLine()
                     report &= "Address Problem: " & vbCrLf
                     report &= "From File: " & anAddrs.ToString & vbCrLf
                     report &= "From db:   " & dbAddress.ToString & vbCrLf & vbCrLf
                  ElseIf String.Compare(anAddrs.Zip, dbAddress.Zip, False) <> 0 Then
                     Console.WriteLine("Address Problem: ")
                     Console.WriteLine("From File: " & anAddrs.ToString)
                     Console.WriteLine("From db:   " & dbAddress.ToString)
                     Console.WriteLine()
                     report &= "Address Problem: " & vbCrLf
                     report &= "From File: " & anAddrs.ToString & vbCrLf
                     report &= "From db:   " & dbAddress.ToString & vbCrLf & vbCrLf
                  End If
               End If
            End If
            sr.Close()
         End Using
      Catch ex As Exception
         Console.WriteLine(ex.Message)
      End Try
   End Sub
   Function inRejectFile(ByVal anOrderNo As String, ByVal anOrderFile As String) As Boolean
      Dim tmpRresult As Boolean = False
      Dim RejectFilePath As String = New FileInfo(anOrderFile).DirectoryName
      Dim RejectFile As String = ""
      For Each f As String In Directory.GetFiles(RejectFilePath)
         If f.ToUpper.Contains("REJECTF") Then
            RejectFile = f
         End If
      Next
      If RejectFile.Length > 3 Then
         If File.Exists(RejectFile) Then
            If File.ReadAllText(RejectFile).ToUpper.Contains("#:" & anOrderNo.ToUpper) Then
               Return True
            End If
         End If
      End If
      Return tmpRresult
   End Function
   Function GetAddressFromDB(ByVal sOrderNo) As OrderAddress
      Dim isMailto As Boolean = False
      Dim tmp As String = ""
      Dim tempCUSTEDP As String = ""
      tempCUSTEDP = GetCUSTEDP(sOrderNo)
      Dim sCUSTEDP As String = ""
      If Not String.IsNullOrEmpty(tempCUSTEDP) Then
         If tempCUSTEDP.ToUpper.StartsWith("OS") Then
            isMailto = True
            sCUSTEDP = tempCUSTEDP.Remove(0, 2)
         Else
            isMailto = False
            sCUSTEDP = tempCUSTEDP
         End If
         Using conn As New SqlConnection(connString1)
            Dim cmd As New SqlCommand("select STREET, CITY, STATE, ZIP from dbo.CUSTOMERS where CUSTEDP = " & sCUSTEDP, conn)
            conn.Open()
            Dim r As SqlDataReader = cmd.ExecuteReader()
            If r.HasRows Then
               Dim tStreet, tCity, tState, tZip As String
               r.Read()
               tStreet = r(0)
               tCity = r(1)
               tState = r(2)
               tZip = r(3)
               r.Close()
               If isMailto Then
                  Return New OrderAddress("S", sOrderNo, "", tStreet.Trim, tCity.Trim, tState.Trim, tZip.Trim)
               Else
                  Return New OrderAddress("C", sOrderNo, "", tStreet.Trim, tCity.Trim, tState.Trim, tZip.Trim)
               End If
            Else
               Return New OrderAddress
            End If
         End Using
      Else
         Return New OrderAddress
      End If
      Return New OrderAddress
   End Function
   Public Function GetCUSTEDP(ByVal ordNo As String) As String
      Dim sCUSTEDP As String = ""
      Using conn As New SqlConnection(connString1)
         Dim cmd As New SqlCommand("select XREFNO from dbo.ORDERXREF where FULLORDERNO like '" & ordNo & "%' and XREFNO like 'OS%'", conn)
         conn.Open()
         Dim r As SqlDataReader = cmd.ExecuteReader()
         If r.HasRows Then
            r.Read()
            sCUSTEDP = r(0)
            r.Close()
            Return sCUSTEDP
         Else
            r.Close()
            cmd = New SqlCommand("select CUSTEDP from dbo.ORDERHEADER where FULLORDERNO like '" & ordNo & "%'", conn)
            r = cmd.ExecuteReader
            If r.HasRows Then
               r.Read()
               sCUSTEDP &= r(0)
               r.Close()
               Return sCUSTEDP
            Else
               Return ""
            End If
         End If
      End Using
      Return "ERROR"
   End Function
   Function ParseCustAddress(ByVal addrline As String) As OrderAddress
      Dim tmpAdd As New OrderAddress
      Dim tmp As String = ""
      Dim orderNo, addr1, sStreet, sCity, sState, sZip As String
      Try
         orderNo = addrline.Substring(2, 8)
         addr1 = addrline.Substring(101, 161 - 101).Trim
         sStreet = addrline.Substring(161, 191 - 161).Trim
         sCity = addrline.Substring(191, 221 - 191).Trim
         sState = addrline.Substring(221, 2).Trim
         sZip = addrline.Substring(223, 237 - 223).Trim
         Return New OrderAddress("C", orderNo, addr1, sStreet, sCity, sState, sZip)
      Catch ex As Exception
         Return New OrderAddress
      End Try
      Return tmpAdd
   End Function
   Function ParseShipToAddress(ByVal addrLine As String) As OrderAddress
      Dim tmp As String = ""
      Dim anAddrs As New OrderAddress
      Dim orderNo, sStreet, sCity, sState, sZip As String
      Try
         orderNo = addrLine.Substring(2, 8)
         sStreet = addrLine.Substring(135, 165 - 135).Trim
         sCity = addrLine.Substring(165, 195 - 165).Trim
         sState = addrLine.Substring(195, 2).Trim
         sZip = addrLine.Substring(197, 210 - 197).Trim
         Return New OrderAddress("S", orderNo, "", sStreet, sCity, sState, sZip)
      Catch ex As Exception
         Return New OrderAddress
      End Try
      Return New OrderAddress
   End Function
   Function GetAddressLines(ByVal aFile As String) As String
      Dim tmp As String = ""
      Try
         Using sr As StreamReader = New StreamReader(aFile)
            Dim aline As String = ""
            Do
               aline = sr.ReadLine()
               If aline.StartsWith("30") Then 'Or aline.StartsWith("30") Then
                  tmp &= aline & vbCrLf
               End If
            Loop Until aline Is Nothing
            sr.Close()
         End Using
      Catch ex As Exception
      End Try
      Return tmp
   End Function
End Module
Public Class OrderAddress
   Private oType As String = ""
   Private oOrderNo As String = ""
   Private oAddress1 As String = ""
   Private oStreet As String = ""
   Private oCity As String = ""
   Private oState As String = ""
   Private oZip As String = ""
   Sub New()
   End Sub
   Sub New(ByVal sOrderNo As String)
      oOrderNo = sOrderNo
   End Sub
   Sub New(ByVal sType As String, ByVal sOrderNo As String, ByVal sAddress1 As String, ByVal sStreet As String, ByVal sCity As String, ByVal sState As String, ByVal sZip As String)
      oType = sType
      oOrderNo = sOrderNo
      oAddress1 = sAddress1
      oStreet = sStreet
      oCity = sCity
      oState = sState
      oZip = sZip
   End Sub
   ReadOnly Property Type() As String
      Get
         Return oType
      End Get
   End Property
   ReadOnly Property OrderNo() As String
      Get
         Return oOrderNo
      End Get
   End Property
   ReadOnly Property Address1() As String
      Get
         Return oAddress1
      End Get
   End Property
   ReadOnly Property Street() As String
      Get
         Return oStreet
      End Get
   End Property
   ReadOnly Property City() As String
      Get
         Return oCity
      End Get
   End Property
   ReadOnly Property State() As String
      Get
         Return oState
      End Get
   End Property
   ReadOnly Property Zip() As String
      Get
         Return oZip
      End Get
   End Property
   Public Overrides Function ToString() As String
      Return String.Format("{0,-2}{1,-10}{2,-30}{3,-35}{4,-30}{5,-4}{6}", oType, oOrderNo, oAddress1, oStreet, oCity, oState, oZip)
   End Function
End Class
