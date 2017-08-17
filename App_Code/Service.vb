Imports System.IO
Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
    Inherits System.Web.Services.WebService
    Public NTTZipFolder As String
    Private tMax(11) As Single
    Private tMin(11) As Single

    <WebMethod()> _
    Public Function DownloadAPEXFolder(NTTFilesFolder As String, session1 As String, type As String) As String
        Dim sFile As String = String.Empty
        Try
            If Not System.IO.Directory.Exists(NTTFilesFolder + session1) Then
                'Return "Folder does not exist"
            End If
            PrepareFiles(NTTFilesFolder + session1)
            zipFolder(NTTFilesFolder.Replace("APEX", "ZIP") + session1, type)

            sFile = NTTZipFolder + "ZIP" + session1 & ".zip"
            Return sFile
            'If IO.File.Exists(sFile) Then
            'End If
            'Return System.IO.File(NTTFilesFolder.Replace("APEX", "ZIP") + session1 + ".zip")
            'showMessage(lblMessage, imgIcon, "Green", "GoIcon.jpg", msgDoc.Descendants("DownloadFolder").Value.Replace("First", "APEX"))
        Catch ex As Exception
            'showMessage(lblMessage, imgIcon, "Red", "StopIcon.jpg", ex.Message)
        Finally
        End Try
    End Function

    <WebMethod()> _
    Public Function GetStates() As DataTable
        Dim sSQL As String = String.Empty
        Dim con As New SqlConnection
        Dim ad As New SqlDataAdapter
        Dim dt As New DataTable

        Try
            con.ConnectionString = dbConnectString("No")
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            sSQL = "SELECT * FROM State ORDER BY [Name]"
            ad.SelectCommand = New SqlCommand(sSQL, con)
            ad.Fill(dt)
            con.Close()
            con.Dispose()
            con = Nothing
            dt.TableName = "States"

            Return (dt)

        Catch ex As Exception
            Return dt
        End Try

    End Function

    <WebMethod()> _
    Public Function GetCounties(state As String) As DataTable
        Dim sSQL As String = String.Empty
        Dim con As New SqlConnection
        Dim ad As New SqlDataAdapter
        Dim dt As New DataTable

        Try
            con.ConnectionString = dbConnectString("No")
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            sSQL = "SELECT * FROM County WHERE StateAbrev like '" & state.Trim & "%' ORDER BY [Name]"
            ad.SelectCommand = New SqlCommand(sSQL, con)
            ad.Fill(dt)
            con.Close()
            con.Dispose()
            con = Nothing
            dt.TableName = "Counties"

            Return (dt)

        Catch ex As Exception
            Return dt
        End Try

    End Function

    <WebMethod()> _
    Public Function GetCounty(state As String, county As String) As DataTable
        Dim sSQL As String = String.Empty
        Dim con As New SqlConnection
        Dim ad As New SqlDataAdapter
        Dim dt As New DataTable

        Try
            con.ConnectionString = dbConnectString("No")
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            sSQL = "SELECT TOP 1 * FROM County_Extended WHERE StateAbrev like '" & state.Trim & "%' AND [Name] like '" & county.Trim & "%' ORDER BY [Name]"
            ad.SelectCommand = New SqlCommand(sSQL, con)
            ad.Fill(dt)
            con.Close()
            con.Dispose()
            con = Nothing
            dt.TableName = "County"

            Return (dt)

        Catch ex As Exception
            Return dt
        End Try

    End Function

    <WebMethod()> _
    Public Function GetRecord(sSql As String) As DataTable
        Dim con As New SqlConnection
        Dim ad As New SqlDataAdapter
        Dim dt As New DataTable

        Try
            con.ConnectionString = dbConnectString("No")
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            'sSQL = "SELECT TOP 1 * FROM County_Extended WHERE StateAbrev like '" & state.Trim & "%' AND [Name] like '" & county.Trim & "%' ORDER BY [Name]"
            ad.SelectCommand = New SqlCommand(sSql, con)
            ad.Fill(dt)
            con.Close()
            con.Dispose()
            con = Nothing
            dt.TableName = "Record"

            Return (dt)

        Catch ex As Exception
            Return dt
        End Try

    End Function

    <WebMethod()> _
    Public Function GetSoilRecord(sSql As String) As DataTable
        Dim con As New SqlConnection
        Dim ad As New SqlDataAdapter
        Dim dt As New DataTable

        Try
            con.ConnectionString = dbConnectString("Soil")
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            'sSQL = "SELECT TOP 1 * FROM County_Extended WHERE StateAbrev like '" & state.Trim & "%' AND [Name] like '" & county.Trim & "%' ORDER BY [Name]"
            ad.SelectCommand = New SqlCommand(sSql, con)
            ad.Fill(dt)
            con.Close()
            con.Dispose()
            con = Nothing
            dt.TableName = "Soil"

            Return (dt)

        Catch ex As Exception
            Return dt
        End Try

    End Function

    <WebMethod()> _
    Public Function GetSSA(county_code As String) As DataTable
        Dim sSQL As String = String.Empty
        Dim con As New SqlConnection
        Dim ad As New SqlDataAdapter
        Dim dt As New DataTable

        Try
            con.ConnectionString = dbConnectString("No")
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            sSQL = "SELECT Code FROM SSArea WHERE CountyCode = '" & county_code & "' ORDER BY [Code]"
            ad.SelectCommand = New SqlCommand(sSQL, con)
            ad.Fill(dt)
            con.Close()
            con.Dispose()
            con = Nothing
            dt.TableName = "SSA"

            Return (dt)

        Catch ex As Exception
            Return dt
        End Try

    End Function

    <WebMethod()> _
    Public Function GetSoils(ssa As String, county_code As String, txtMaxSlope As String, txtSoilsPercentage As String) As DataTable
        Dim sSQL As String = String.Empty
        Dim con As New SqlConnection
        Dim ad As New SqlDataAdapter
        Dim dt As New DataTable

        Try
            con.ConnectionString = dbConnectString("Soil")
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If
            Dim numericCheck As Boolean = IsNumeric(txtSoilsPercentage)
            If txtSoilsPercentage = 100 Or txtSoilsPercentage <= 0 Or numericCheck = False Then
                sSQL = "SELECT * FROM " & county_code.Substring(0, 2) & "SOILS WHERE TSSSACode = '" & ssa & "' AND (slopel+slopeh)/2 <= " & txtMaxSlope & "  ORDER BY [Series], [SeriesName], [ldep]"
            Else
                sSQL = "SELECT TOP " & txtSoilsPercentage & " PERCENT * FROM " & county_code.Substring(0, 2) & "SOILS WHERE TSSSACode = '" & ssa & "' AND (slopel+slopeh)/2 <= " & txtMaxSlope & "  ORDER BY horizdesc1, [Series], [SeriesName], [ldep]"
            End If
            ad.SelectCommand = New SqlCommand(sSQL, con)
            ad.Fill(dt)
            con.Close()
            con.Dispose()
            con = Nothing
            dt.TableName = "Soils"

            Return (dt)

        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Private Sub PrepareFiles(folderPath As String)
        Dim file As String
        Dim folderZip As String = folderPath.Replace("APEX", "ZIP")
        Dim APEXFiles As System.IO.StreamReader
        Dim folder As String = My.Computer.FileSystem.GetParentPath(Server.MapPath("")) + "\weather\"
        APEXFiles = New System.IO.StreamReader(folder & "/Resources/APEX1Files.txt")
        System.IO.Directory.CreateDirectory(folderZip)

        Do While APEXFiles.EndOfStream <> True
            file = APEXFiles.ReadLine()
            If file.ToLower.Contains("apex") Or file.ToLower.Contains(".wnd") Or file.ToLower.Contains(".wp1") Or file.ToLower.Contains(".dat") Then
                If System.IO.File.Exists(folderPath & "\" & file.Trim) Then
                    System.IO.File.Copy(folderPath & "\" & file.Trim, folderZip & "\" & file.Trim, True)
                End If
            End If
        Loop

        Dim files As New IO.DirectoryInfo(folderPath)
        For Each file1 As IO.FileInfo In files.GetFiles("*.wp1", IO.SearchOption.AllDirectories)
            System.IO.File.Copy(folderPath & "\" & file1.Name, folderZip & "\" & file1.Name, True)
        Next

        For Each file1 As IO.FileInfo In files.GetFiles("*.wnd", IO.SearchOption.AllDirectories)
            System.IO.File.Copy(folderPath & "\" & file1.Name, folderZip & "\" & file1.Name, True)
        Next

        For Each file1 As IO.FileInfo In files.GetFiles("*.opc", IO.SearchOption.AllDirectories)
            System.IO.File.Copy(folderPath & "\" & file1.Name, folderZip & "\" & file1.Name, True)
        Next

        For Each file1 As IO.FileInfo In files.GetFiles("*.sol", IO.SearchOption.AllDirectories)
            System.IO.File.Copy(folderPath & "\" & file1.Name, folderZip & "\" & file1.Name, True)
        Next
    End Sub

    Public Function zipFolder(zipPath As String, type As String) As String
        Try
            If System.Net.Dns.GetHostName = "T-NN" Then
                NTTZipFolder = "E" & ":\NTTSLZip\"
            Else
                NTTZipFolder = "C" & ":\NTTSLZip\"
            End If

            Using zip As New Ionic.Zip.ZipFile
                zip.UseUnicodeAsNecessary = True
                zip.CompressionLevel = Ionic.Zlib.CompressionLevel.BestCompression
                zip.Comment = "This zip was created at " & System.DateTime.Now.ToString("G")
                Select Case type.ToLower
                    Case "dndc"
                        zip.AddDirectory(zipPath & "\DnDc")
                        zip.Save(NTTZipFolder & zipPath.Split("/")(2) & ".zip")
                    Case "apex"
                        zip.AddDirectory(zipPath)
                        zip.Save(NTTZipFolder & zipPath.Split("/")(2) & ".zip")
                End Select

            End Using
            Return "OK"
        Catch ex As System.Exception
            Return ex.Message
        End Try
    End Function

    <WebMethod()> _
    Public Function GetWeather(path As String) As System.Collections.Generic.List(Of String)
        Dim sr As System.IO.StreamReader = System.IO.File.OpenText(path)
        Dim weather As New System.Collections.Generic.List(Of String)
        Dim temp As String = String.Empty

        Do While Not sr.EndOfStream
            weather.Add(sr.ReadLine)
        Loop
        sr.Close()
        sr.Dispose()
        sr = Nothing
        Return weather
    End Function

    <WebMethod()> _
    Public Function Create_wp1_from_weather(loc As String, wp1name As String, controlvalue5 As Integer, pgm As String) As String()
        Dim sr As StreamReader = New StreamReader(loc & "\APEX.wth")
        Dim sw As StreamWriter = New StreamWriter(loc & "\" & wp1name.Trim & ".tmp")
        File.Copy("E:\Weather\wp1File\" & wp1name.Trim & ".wp1", loc & "\" & wp1name.Trim & ".wp1", True)
        Dim sr1 As StreamReader = New StreamReader(loc & "\" & wp1name.Trim & ".wp1")
        Dim wthData As WthData
        Dim wthDatas As New List(Of WthData)
        Dim wthMonthData As WthData
        Dim wthMonthDatas As New List(Of WthData)
        Dim temp As String = String.Empty
        Dim maxTemp As Single = 0
        Dim minTemp As Single = 0
        Dim pcp As Single = 0
        Dim solarR As Single = 0
        Dim relativeH As Single = 0
        Dim windS As Single = 0
        Dim month As UShort
        Dim year As UShort
        Dim day As UShort
        Dim wp1Data As Wp1Data
        Dim wp1Datas As New List(Of Wp1Data)
        Dim wp1MonSD As New Wp1Data
        Dim wp1MonSDs As New List(Of Wp1Data)
        Dim monthAnt As UShort = 0
        Dim firstYear As UShort = 0
        Dim lastYear As UShort = 0
        Dim newMonth As Boolean = False
        Dim dry_day_ant As Boolean = False
        Dim lines(15) As String

        Try
            Do While sr.EndOfStream <> True
                temp = sr.ReadLine
                If temp.Trim = "" Then
                    Exit Do
                End If
                UShort.TryParse(temp.Substring(2, 4), year)
                UShort.TryParse(temp.Substring(6, 4), month)
                UShort.TryParse(temp.Substring(10, 4), day)
                Single.TryParse(temp.Substring(14, 6), solarR)
                Single.TryParse(temp.Substring(20, 6), maxTemp)
                Single.TryParse(temp.Substring(26, 6), minTemp)
                If temp.Length < 39 Then
                    Single.TryParse(temp.Substring(32, 6), pcp)
                Else
                    Single.TryParse(temp.Substring(32, 7), pcp)
                End If
                If temp.Length >= 49 Then
                    Single.TryParse(temp.Substring(44, 6), windS)
                End If
                If temp.Length >= 43 Then
                    Single.TryParse(temp.Substring(38, 6), relativeH)
                End If
                wthData = New WthData
                wthData.Day = day
                wthData.Month = month
                wthData.Year = year
                If wp1Datas.Count < month Then
                    wp1Data = New Wp1Data
                    newMonth = True
                    wp1Datas.Add(wp1Data)
                End If
                If solarR > -900 And solarR < 900 Then wp1Datas(month - 1).Obsl += solarR : wp1Datas(month - 1).Days_obsl += 1
                wthData.MaxTemp = maxTemp
                If maxTemp > -900 And maxTemp < 900 Then wp1Datas(month - 1).Obmx += maxTemp : wp1Datas(month - 1).Days_obmx += 1
                wthData.MinTemp = minTemp
                If minTemp > -900 And minTemp < 900 Then wp1Datas(month - 1).Obmn += minTemp : wp1Datas(month - 1).Days_obmn += 1
                wthData.Pcp = pcp
                If pcp > -900 And pcp < 900 Then wp1Datas(month - 1).Rmo += pcp : wp1Datas(month - 1).Days_rmo += 1
                wp1Datas(month - 1).Rmosd += pcp
                wp1Datas(month - 1).Rh += relativeH
                wp1Datas(month - 1).Uav0 += windS
                wthDatas.Add(wthData)
                If monthAnt <> month Then
                    If monthAnt <> 0 Then
                        wthMonthData.MaxTemp /= wthMonthData.Day
                        wthMonthData.MinTemp /= wthMonthData.Day
                        wthMonthData.Pcp /= wthMonthData.Day
                        wthMonthDatas.Add(wthMonthData)
                        'calculate SD for each month and add to wp1datas
                        wthMonthData.MaxTemp = Math.Sqrt(wp1MonSDs.Where(Function(x) x.Obmx < 999).Sum(Function(x) (wthMonthData.MaxTemp - x.Obmx) ^ 2) / wp1MonSDs.Where(Function(x) x.Obmx < 999).Count)
                        wthMonthData.MinTemp = Math.Sqrt(wp1MonSDs.Where(Function(x) x.Obmn < 999).Sum(Function(x) (wthMonthData.MinTemp - x.Obmn) ^ 2) / wp1MonSDs.Where(Function(x) x.Obmn < 999).Count)
                        wp1MonSDs.Clear()
                    Else
                        firstYear = year
                    End If
                    monthAnt = month
                    wthMonthData = New WthData
                    wthMonthData.Year = year
                    wthMonthData.Month = month
                End If
                wp1MonSD = New Wp1Data
                wthMonthData.Day += 1
                wthMonthData.MaxTemp += maxTemp
                wp1MonSD.Obmx = maxTemp
                wthMonthData.MinTemp += minTemp
                wp1MonSD.Obmn = minTemp
                wthMonthData.Pcp += pcp
                wp1MonSD.Rmo = pcp
                wp1MonSDs.Add(wp1MonSD)
                If pcp > 0 Then
                    wthMonthData.WetDay += 1
                    If dry_day_ant = True Then wthMonthData.dd_wd += 1
                    dry_day_ant = False
                Else
                    dry_day_ant = True
                End If
            Loop
            'calculate the number of years
            Dim years = year - firstYear + 1
            'calculate averages for each month
            For Each mon In wp1Datas
                mon.Obmx /= mon.Days_obmx
                mon.Obmn /= mon.Days_obmn
                mon.Rmo /= years
                mon.Rmosd /= mon.Days_rmo
                mon.Obsl /= mon.Days_obsl
                If mon.Days_rh = 0 Then mon.Rh = 0 Else mon.Rh /= mon.Days_rh
                If mon.Days_Uav0 = 0 Then mon.Uav0 = 0 Else mon.Uav0 /= mon.Days_Uav0
                mon.Wi = 0
            Next
            '************************add the last month of the last year*********************
            wthMonthData.MaxTemp /= wthMonthData.Day
            wthMonthData.MinTemp /= wthMonthData.Day
            wthMonthData.Pcp /= wthMonthData.Day
            If pcp > 0 Then wthMonthData.WetDay += 1
            wthMonthDatas.Add(wthMonthData)
            'calculate SD for each month and add to wp1datas
            wthMonthData.MaxTemp = Math.Sqrt(wp1MonSDs.Where(Function(x) x.Obmx < 999).Sum(Function(x) (wthMonthData.MaxTemp - x.Obmx) ^ 2) / wp1MonSDs.Where(Function(x) x.Obmx < 999).Count)
            wthMonthData.MinTemp = Math.Sqrt(wp1MonSDs.Where(Function(x) x.Obmn < 999).Sum(Function(x) (wthMonthData.MinTemp - x.Obmn) ^ 2) / wp1MonSDs.Where(Function(x) x.Obmn < 999).Count)
            wp1MonSDs.Clear()
            '********************************************************************************
            'calculate total days per month in the whole period
            Dim day_30 As UShort = years * 30
            Dim day_31 As UShort = years * 31
            Dim day_Feb As UShort = years * 28 + years \ 4
            Dim pwd As Single = 0 'probability of wet day
            Dim b1 = 0.75
            Dim numerator, denominator As Single

            For i = 1 To 12
                month = i
                'calculate b1 
                b1 = wthMonthDatas.Where(Function(x) x.Month = month).Sum(Function(x) x.dd_wd) / wthMonthDatas.Where(Function(x) x.Month = month).Sum(Function(x) x.WetDay)
                wp1Datas(i - 1).Sdtmx = wthMonthDatas.Where(Function(x) x.Month = month).Average(Function(x) x.MaxTemp)
                wp1Datas(i - 1).Sdtmn = wthMonthDatas.Where(Function(x) x.Month = month).Average(Function(x) x.MinTemp)
                wp1Datas(i - 1).Rst2 = Math.Sqrt(wthDatas.Where(Function(x) x.Month = month And x.Pcp < 999).Sum(Function(x) (wp1Datas(month - 1).Rmosd - x.Pcp) ^ 2) / wthDatas.Where(Function(x) x.Month = month And x.Pcp < 999).Count)
                numerator = wthDatas.Where(Function(x) x.Month = month And x.Pcp < 999).Sum(Function(x) (x.Pcp - wp1Datas(month - 1).Rmosd) ^ 3) / wthDatas.Where(Function(x) x.Month = month And x.Pcp < 999).Count
                denominator = (wthDatas.Where(Function(x) x.Month = month And x.Pcp < 999).Sum(Function(x) (x.Pcp - wp1Datas(month - 1).Rmosd) ^ 2) / (wthDatas.Where(Function(x) x.Month = month And x.Pcp < 999).Count - 1)) ^ (3 / 2)
                wp1Datas(i - 1).Rst3 = numerator / denominator
                wp1Datas(i - 1).Uavm = wthMonthDatas.Where(Function(x) x.Month = month).Average(Function(x) x.WetDay)
                Select Case month
                    Case 1, 3, 5, 7, 8, 10, 12
                        pwd = wthDatas.Where(Function(x) x.Month = month And x.Pcp < 999 And x.Pcp > 0).Count / day_31
                    Case 2
                        pwd = wthDatas.Where(Function(x) x.Month = month And x.Pcp < 999 And x.Pcp > 0).Count / day_Feb
                    Case 4, 6, 9, 11
                        pwd = wthDatas.Where(Function(x) x.Month = month And x.Pcp < 999 And x.Pcp > 0).Count / day_30
                End Select
                wp1Datas(i - 1).Prw1 = b1 * pwd   'taking from http://www.nrcs.usda.gov/Internet/FSE_DOCUMENTS/nrcs143_013182.pdf page 5
                wp1Datas(i - 1).Prw2 = 1.0 - b1 + wp1Data.Prw1   'taking from http://www.nrcs.usda.gov/Internet/FSE_DOCUMENTS/nrcs143_013182.pdf page 5
            Next

            'titles
            Dim j As Short = 0
            Do While sr1.EndOfStream = False Or j < 16
                lines(j) = sr1.ReadLine
                j += 1
            Loop
            j = 0
            Dim no_rh As Boolean = False
            For Each wp1 In wp1Datas
                If j = 0 Then
                    lines(2) = Math.Round(wp1.Obmx, 2).ToString("N2").PadLeft(6)
                    lines(3) = Math.Round(wp1.Obmn, 2).ToString("N2").PadLeft(6)
                    lines(4) = Math.Round(wp1.Sdtmx, 2).ToString("N2").PadLeft(6)
                    lines(5) = Math.Round(wp1.Sdtmn, 2).ToString("N2").PadLeft(6)
                    lines(6) = Math.Round(wp1.Rmo, 1).ToString("N1").PadLeft(6)
                    lines(7) = Math.Round(wp1.Rst2, 1).ToString("N1").PadLeft(6)
                    lines(8) = Math.Round(wp1.Rst3, 2).ToString("N2").PadLeft(6)
                    lines(9) = Math.Round(wp1.Prw1, 3).ToString("N3").PadLeft(6)
                    lines(10) = Math.Round(wp1.Prw2, 3).ToString("N3").PadLeft(6)
                    lines(11) = Math.Round(wp1.Uavm, 2).ToString("N2").PadLeft(6)
                    lines(12) = Math.Round(0, 2).ToString("N2").PadLeft(6)
                    lines(13) = Math.Round(wp1.Obsl, 2).ToString("N2").PadLeft(6)
                    If pgm = "APEX" Then
                        lines(14) = Math.Round(wp1.Rh, 2).ToString("N2").PadLeft(6) : no_rh = True
                    Else
                        If wp1.Rh > 0 Then lines(14) = Math.Round(wp1.Rh, 2).ToString("N2").PadLeft(6) : no_rh = True
                    End If
                    lines(15) = Math.Round(wp1.Uav0, 2).ToString("N2").PadLeft(6)
                    j = 1
                Else
                    lines(2) &= Math.Round(wp1.Obmx, 2).ToString("N2").PadLeft(6)
                    lines(3) &= Math.Round(wp1.Obmn, 2).ToString("N2").PadLeft(6)
                    lines(4) &= Math.Round(wp1.Sdtmx, 2).ToString("N2").PadLeft(6)
                    lines(5) &= Math.Round(wp1.Sdtmn, 2).ToString("N2").PadLeft(6)
                    lines(6) &= Math.Round(wp1.Rmo, 1).ToString("N1").PadLeft(6)
                    lines(7) &= Math.Round(wp1.Rst2, 1).ToString("N1").PadLeft(6)
                    lines(8) &= Math.Round(wp1.Rst3, 2).ToString("N2").PadLeft(6)
                    lines(9) &= Math.Round(wp1.Prw1, 3).ToString("N3").PadLeft(6)
                    lines(10) &= Math.Round(wp1.Prw2, 3).ToString("N3").PadLeft(6)
                    lines(11) &= Math.Round(wp1.Uavm, 2).ToString("N2").PadLeft(6)
                    lines(12) &= Math.Round(0, 2).ToString("N2").PadLeft(6)
                    lines(13) &= Math.Round(wp1.Obsl, 2).ToString("N2").PadLeft(6)
                    If no_rh Then lines(14) &= Math.Round(wp1.Rh, 2).ToString("N2").PadLeft(6)
                    lines(15) &= Math.Round(wp1.Uav0, 2).ToString("N2").PadLeft(6)
                End If
            Next
            For line = 0 To 15
                'For Each line In lines. If lines 14-16 has information in wth file it will be calculated if not it will be zeros.
                'SR, RH, and Wind Speed. Line 13 is always zeros.
                sw.WriteLine(lines(line))
            Next
            If Not sr Is Nothing Then
                sr.Close()
                sr.Dispose()
                sr = Nothing
            End If
            If Not sr1 Is Nothing Then
                sr1.Close()
                sr1.Dispose()
                sr1 = Nothing
            End If
            If Not sw Is Nothing Then
                sw.Close()
                sw.Dispose()
                sw = Nothing
            End If

            Return lines

            File.Copy(loc & "\" & wp1name.Trim & ".wp1", loc & "\" & wp1name.Trim & ".org", True)
            File.Copy(loc & "\" & wp1name.Trim & ".tmp", loc & "\" & wp1name.Trim & ".wp1", True)
        Catch ex As Exception
            Dim msg As String = ex.Message
        Finally
            If Not sr Is Nothing Then
                sr.Close()
                sr.Dispose()
                sr = Nothing
            End If
            If Not sr1 Is Nothing Then
                sr1.Close()
                sr1.Dispose()
                sr1 = Nothing
            End If
            If Not sw Is Nothing Then
                sw.Close()
                sw.Dispose()
                sw = Nothing
            End If
        End Try
    End Function

    <WebMethod()> _
    Public Function getHU(ByVal crop As Short, ByVal nlat As Single, nlon As Single, plantingJulianDay As UShort, path As String) As Single
        Dim srCrop As StreamReader = New StreamReader("C:\NTTDeployed\NTTWebService\calcHU\PHUCRP.DAT")
        'Dim srWp1 As StreamReader
        'Dim tMax(11) As Single
        'Dim tMin(11) As Single
        Dim wp1aLat As Single
        Dim ivar(300), cNum, jPlace As Short
        Dim j, phuc, nt As Single
        Dim mo, hd(300) As Single
        Dim itil As Short = 0
        Dim cname As String = String.Empty
        Dim to1 As Single = 0 : Dim dd As Single = 0 : Dim daym As Single = 0 : Dim iPlant As Single = 0 : Dim tb As Single = 0
        Dim phu As Single = 0 : Dim phus As Single = 0
        Dim temp As String = String.Empty
        Dim ta As Single = 0
        Dim sep As String = " "
        Dim phux As Single
        Dim ida As Short = 0
        Dim k As Short = 0
        Dim ncc() As Short = {0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366}
        Dim nc() As Short = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31}
        Dim offset As Integer = 6
        'Const latDif As Single = 1.0       'initial value 0.04
        'Const lonDif As Single = 1.0       'initial value 0.09
        'Dim latLess, latPlus, lonLess, lonPlus As Double
        Dim wp1Name As String = String.Empty

        Try
            GetWeatherFile(nlat, nlon, path)   'used just to create the Tmax and tMin from the real weather information.
            'latLess = nlat - latDif
            'latPlus = nlat + latDif
            'lonLess = nlon - lonDif
            'lonPlus = nlon + lonDif
            'wp1Name = GetWp1Name(nlat, nlon, latLess, latPlus, lonLess, lonPlus)
            'srWp1 = New StreamReader("E:\Weather\wp1File\" & wp1Name)
            nt = 1
            phuc = 0
            Dim i As Short
            For i = 0 To 299
                ivar(i) = 0
            Next

            jPlace = 0
            nt = 1

            j = 0

            temp = srCrop.ReadLine
            Do While Not srCrop.EndOfStream()
                cNum = temp.Substring(0, 3)
                itil = temp.Substring(12, 4)
                to1 = temp.Substring(16, 8)
                tb = temp.Substring(24, 8)
                dd = temp.Substring(32, 8)
                daym = temp.Substring(40, 8)
                If cNum = crop Then Exit Do
                j += 1
                temp = srCrop.ReadLine
            Loop
            srCrop.Close()
            srCrop.Dispose()
            srCrop = Nothing
            '/************ THIS BLOCK CHANGED TO TAKE THE NUMBERS FROM THE CURRENT WEAHTER FILE.  *****/
            'srWp1.ReadLine()
            'wp1aLat = srWp1.ReadLine.Substring(8, 7).Trim
            'temp = srWp1.ReadLine
            'For i = 0 To 11
            '    tMax(i) = Val(temp.Substring(offset * i, 6))
            'Next
            'temp = srWp1.ReadLine
            'For i = 0 To 11
            '    tMin(i) = Val(temp.Substring(offset * i, 6))
            'Next
            'srWp1.Close()
            'srWp1.Dispose()
            'srWp1 = Nothing
            '/*********************************************************************************************
            '/****** THIS IS THE NEW BLOCK ************************************************************/
            wp1aLat = nlat          'LATITUDE IS BEING PASSED   

            '/********************************************************************************************

            jPlace = j
            '*****varieties
            phux = 30 * 9.5 / ((wp1aLat + 1) ^ 0.5)
            '*****calculate plant date for summer crop without plant date input
            If daym > 0 Then
                phuc = 0
                For mo = 0 To 11
                    For ida = 0 To nc(mo) - 1
                        ta = (tMax(mo) + tMin(mo)) / 2 - (tb + 3)
                        If ta > 0 Then
                            phuc = phuc + ta
                            If phuc > phux Then GoTo 50
                        End If
                    Next
                Next
50:             iPlant = jdt(ida + 1, mo + 1, nt)
                k = iPlant + daym

                If plantingJulianDay > 0 Then
                    k = plantingJulianDay + daym
                    If plantingJulianDay + daym > 365 Then k = 365 Else k = plantingJulianDay + daym
                    phu = ahu(plantingJulianDay, k, mo + 1, tMax, tMin, tb, to1)
                    If plantingJulianDay + daym - 365 > 0 Then
                        phu = ahu(1, plantingJulianDay + daym - 365, 1, tMax, tMin, tb, to1) + phu
                    End If
                    Return phu
                End If
                '*****iplant=julian day of planting
                '*****daym=crop maturity days
                phu = ahu(iPlant, k, mo + 1, tMax, tMin, tb, to1)
                If (daym <= 0) Then phu = 0
                '*****determine planting date for winter crops
                If (itil = 2) Then
                    phuc = 0
                    phux = phu / 2
                    Dim mo1 As Short
                    For mo = 0 To 11
                        mo1 = 11 - mo
                        For ida = 0 To nc(mo1)
                            ta = (tMax(mo1) + tMin(mo1)) / 2 - tb
                            If (ta > 0) Then phuc = phuc + ta
                            If (phuc > phux) Then GoTo 80
                        Next
                    Next
80:                 ida = nc(mo1) - ida
                    iPlant = jdt(ida + 1, mo1 + 1, nt)
                End If
            End If
            mo = xmonth(iPlant, mo)
            '*****select the proper variety
            '     sum heat units from plant to end of year
            If (phu <= 0) Then GoTo 130
            If (itil = 1) Then
                phuc = ahu(iPlant, 365, mo + 1, tMax, tMin, tb, to1)
            Else
                phuc = ahu(iPlant, 365, mo + 1, tMax, tMin, tb, to1)
                If iPlant + daym - 365 > 0 Then
                    phuc = ahu(1, iPlant + daym - 365, 1, tMax, tMin, tb, to1) + phuc
                End If
            End If
            '****changed constant 1.2 to input variable dd(*)+1.    02/23/95 nbs
            hd(j) = 1 + (dd * ((1 - ((wp1aLat + 1) ^ 0.5 / 9.5))) ^ 0.1)
            phuc = phuc / hd(j)
            If (phuc > phu) Then GoTo 120

            '120:            If (ic > 1) Then ic = 1  'ic con not be > 1 because ic went form 1 to 1 
120:        If (phu <= 0) Then
                ivar(j) = 0
            Else
                ivar(j) = 1
            End If
130:        If (ivar(j) = 0 And phu > 0) Then ivar(j) = (phuc * hd(j) / phu) * 100 'this condition never is true because if phu is > 0 ivar=1
            j = 1
            '*****overrides plant optimal temperature in upper limit on heat units
            '*****accumulated on a particular day
            '*****this just simulates the total heat units per year for base 0.
            to1 = 150
            tb = 0
            Dim phutot As Single
            phutot = ahu(1, 365, mo + 1, tMax, tMin, tb, to1)
            j = jPlace
            to1 = 150
            tb = 0
            phus = ahu(1, iPlant, mo + 1, tMax, tMin, tb, to1) / phutot

            If (itil = 1) Then
                Return phu
            Else
                Return phuc
            End If
        Catch ex As System.Exception
            Return "Error 407 - " & ex.Message & " - Error calculating Heat Units"
        End Try
    End Function

    Function dbConnectString(db As String) As String  'db = "No correspond to NTTDB tables; Soil corespond to Soil database.
        Dim sConnectString As String = String.Empty

        'If System.Net.Dns.GetHostName = "T-NN" Then
        '    sConnectString = "Data Source=T-NN\SQLEXPRESS;Initial Catalog=NTTDB;Persist Security Info=True;User ID=sa;Password=pass$word"
        'Else
        '    sConnectString = "Data Source=104.239.136.28;Initial Catalog=NTTDB;Persist Security Info=True;User ID=sa;Password=pass$word"
        'End If
        If db = "Soil" Then
            sConnectString = "Data Source=T-NN1\SQLEXPRESS;Initial Catalog=SSURGOSOILDB2014;Persist Security Info=True;User ID=sa;Password=pass$word"
        Else
            sConnectString = "Data Source=T-NN\SQLEXPRESS;Initial Catalog=NTTDB;Persist Security Info=True;User ID=sa;Password=pass$word"
        End If
        Return sConnectString
    End Function

    Function ahu(ByVal m As Single, ByVal k As Short, ByVal mo As Short, ByVal tmax() As Single, ByVal tmin() As Single, ByVal tbj As Single, ByVal to1j As Single) As Single
        Dim ahu1 As Single = 0
        Dim l As Short
        Dim ta As Single
        Dim humx As Single
        '     this subroutine accumulates heat units and radiation to calculate
        '     the potential heat units
        For l = m To k
            mo = xmonth(l, mo)
            ta = (tmax(mo - 1) + tmin(mo - 1)) / 2 - tbj
            If (ta > 0) Then
                humx = to1j - tbj
                If (ta > humx) Then ta = humx
                ahu1 = ahu1 + ta
            End If
        Next

        Return ahu1
    End Function

    Function xmonth(ByRef ida As Short, ByRef Month As Short) As Short
        Dim mday As Short
        '     this subroutine determines the month, given the day of the year
        '     +++ ARGUMENTS +++
        '     ida - integer day of the year
        '     month - integer month of the year
        If (ida < 32) Then
            Month = 1
            mday = ida
        ElseIf (ida < 61) Then
            Month = 2
            mday = ida - 31
        ElseIf (ida < 92) Then
            Month = 3
            mday = ida - 60
        ElseIf (ida < 122) Then
            Month = 4
            mday = ida - 91
        ElseIf (ida < 153) Then
            Month = 5
            mday = ida - 121
        ElseIf (ida < 183) Then
            Month = 6
            mday = ida - 152
        ElseIf (ida < 214) Then
            Month = 7
            mday = ida - 182
        ElseIf (ida < 245) Then
            Month = 8
            mday = ida - 213
        ElseIf (ida < 275) Then
            Month = 9
            mday = ida - 244
        ElseIf (ida < 306) Then
            Month = 10
            mday = ida - 274
        ElseIf (ida < 336) Then
            Month = 11
            mday = ida - 305
        Else
            Month = 12
            mday = ida - 335
        End If
        Return Month
    End Function

    <WebMethod()> _
    Public Function jdt(ByVal i As Short, ByVal m As Short, ByVal nt As Single) As Short
        '     THIS SUBROUTINE COMPUTES THE DAY OF THE YEAR, GIVEN THE MONTH AND
        '     THE DAY OF THE MONTH.
        Dim nb() As Short = {0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366}
        Dim jdt1 As Short = 0
        'nb = 
        If (m <> 0) Then
            If (m <= 2) Then
                jdt1 = nb(m - 1) + i
                Return jdt1
            End If
            jdt1 = nb(m - 1) - nt + i
        Else
            jdt1 = 0
        End If
        Return jdt1
    End Function

    Public Function GetWeatherFile(nlat As Single, nlon As Single, path As String) As System.Collections.Generic.List(Of String)
        'Dim path As String = String.Empty
        Dim weather As New System.Collections.Generic.List(Of String)
        Dim temp As String = String.Empty
        Const latDif As Single = 0.04
        Const lonDif As Single = 0.09
        Dim latLess, latPlus, lonLess, lonPlus As Double

        latLess = nlat - latDif : latPlus = nlat + latDif
        lonLess = nlon - lonDif : lonPlus = nlon + lonDif
        'this was changed to take the weather file already careated in the APEX file.
        path = path & "\APEX.wth"
        'path = "E:\Weather\weatherFiles\US\" & GetWeatherfileName(nlat, nlon, latLess, latPlus, lonLess, lonPlus)
        Dim weatherList As New List(Of Weather)
        If Not path.Contains("Error") Then
            Dim sr As System.IO.StreamReader = System.IO.File.OpenText(path)
            Dim i As Integer = 0
            Dim weatherRow As Weather

            Do While Not sr.EndOfStream
                weather.Add(sr.ReadLine)
                weatherRow = New Weather
                weatherRow.Year = weather.Item(i).Substring(2, 4)
                weatherRow.Month = weather.Item(i).Substring(6, 4)
                weatherRow.Day = weather.Item(i).Substring(10, 4)
                weatherRow.Srad = weather.Item(i).Substring(14, 6)
                weatherRow.Tmax = weather.Item(i).Substring(20, 6)
                weatherRow.Tmin = weather.Item(i).Substring(26, 6)
                weatherRow.Prcp = weather.Item(i).Substring(32, 6)
                weatherList.Add(weatherRow)
                i += 1
            Loop
            sr.Close()
            sr.Dispose()
            sr = Nothing

            For i = 0 To 11
                tMax(i) = weatherList.Where(Function(x) x.Month = i + 1).Average(Function(x) x.Tmax)
                tMin(i) = weatherList.Where(Function(x) x.Month = i + 1).Average(Function(x) x.Tmin)
            Next
        End If
        Return weather
    End Function

    <WebMethod()> _
    Public Function GetWeatherfileName(ByVal nlat As Single, ByVal nlon As Single, ByVal latLess As Single, ByVal latPlus As Single, ByVal lonLess As Single, ByVal lonPlus As Single) As DataTable
        Dim sSQL As String = String.Empty
        'Dim dr As SqlDataReader = Nothing
        Dim dt As New DataTable
        Dim con As New SqlConnection
        Dim ad As New SqlDataAdapter
        Dim wName As String = String.Empty

        Try
            con.ConnectionString = dbConnectString("No")
            If con.State = ConnectionState.Closed Then
                con.Open()
            End If

            sSQL = "SELECT TOP 1 FileName, initialYear, finalYear, (lat - " & nlat & ") + (lon + " & nlon & ") as distance FROM weatherCoor " & _
                   "WHERE lat > " & latLess & " and lat < " & latPlus & " and lon > " & lonLess & " and lon < " & lonPlus
            ad.SelectCommand = New SqlCommand(sSQL, con)
            ad.Fill(dt)
            con.Close()
            con.Dispose()
            con = Nothing
            dt.TableName = "WeatherFile"

            'wName = dt.Rows(0).Item("FileName")

            Return dt

        Catch ex As Exception
            Return dt
        End Try
    End Function
End Class