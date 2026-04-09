Imports System.Net
Imports System.Xml
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium.Support.UI
Imports System.IO
Imports System.Threading
Imports WebDriverManager
Imports WebDriverManager.DriverConfigs.Impl
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel


'Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports SeleniumExtras.WaitHelpers

Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Canvas.Parser



Module Module1

    Sub Main()


        Dim currentUser As String = Environment.UserName.ToLower()

        ' === WHITELIST USER ===
        Dim allowedUsers As String() = {
    "ingrid.clara",
    "albert.shho",
    "pantaleon.carnegie",
    "sitti.sabdan",
    "tiarni.shbn",
    "maria.wasonono",
    "nyoman.pandiarta",
    "soni.candra",
    "zulkifli.ryu",
    "komangayu.wijayanti",
    "agus.aryana",
    "wayan.lawinataliani",
    "pajak.shasri",
    "rustanty.sibarani",
    "christie.laurensia",
    "grady.gunawan",
    "andri.andri",
    "elvi.elvi",
    "dea.nasution",
    "seyfia.fitri",
    "dewi.maribeth",
    "yessy.sihotang",
    "indri.wulandari",
    "claudia.meke",
    "novina.kosasih",
    "accounting.shbp",
    "meilisa.zefania",
    "putri.gratia",
    "yogi.wijaya",
    "ariva.puspa",
    "ardy.susanto",
    "nizar.adriadi",
    "elisa.hebingadil",
    "vanny.tilaar",
    "agatha.fabyan",
    "wike.widyawati",
    "lusinda.sagala",
    "duffin.duffin",
    "andrias.prastowo",
    "iing.kurniawan",
    "jen.rimbun",
    "andri.napitu",
    "mira.kristy",
    "astrid.vellyanie",
    "franca.irmadita",
    "birgitta.hardanti",
    "maryati.situmorang",
    "theodora.lesilolo",
    "antonius.maturbongs",
    "steven.setiawan",
    "mita.irawati",
    "risma.maria",
    "rosalia.lewoleba",
    "angela.marici",
    "iskandar.petrus",
    "angelina.nurdin",
    "sanjaya",
    "andri.kurniawan",
    "susi.damayanti",
    "leo.marden",
    "frisca.napitupulu",
    "devi.gustina",
    "rikky.irawan",
    "johan.tanuwijaya",
    "reggy.natanael",
    "irhandy.pisenli",
    "catherine.eduardus",
    "nangsi.siagian",
    "krisna.sriloka",
    "agnes.irene",
    "vivin.elvira"
}



        Dim isAllowed As Boolean = allowedUsers.Contains(currentUser)

        If Not isAllowed Then
            MsgBox(
        "Akses ditolak." & vbCrLf &
        "User: " & currentUser & vbCrLf &
        "Tidak terdaftar di sistem.",
        MsgBoxStyle.Critical,
        "Security Check"
    )
            Exit Sub
        End If


        Console.WriteLine("=================================")
        Console.WriteLine(" PILIH MODE PROSES ")
        Console.WriteLine("=================================")
        Console.WriteLine("1. Login & Download PDF dari Coretax")
        Console.WriteLine("2. Insert PDF ke Excel")
        Console.WriteLine("0. Keluar")
        Console.WriteLine("---------------------------------")
        Console.Write("Pilihan Anda: ")

        Dim pilihan As String = Console.ReadLine().Trim()

        Select Case pilihan
            Case "1"
                ProsesDownloadPDF()
            Case "2"
                ProsesPdfKeExcel()
            Case "0"
                Exit Sub
            Case Else
                Console.WriteLine("Pilihan tidak valid.")
        End Select

        Console.WriteLine()
        Console.WriteLine("Klik Close untuk keluar...")
        Console.ReadLine()


    End Sub




    Sub ProsesDownloadPDF()

        ' ==============================
        ' SETUP LOG FILE
        ' ==============================
        Dim logPath As String = ""

        Dim jawab As MsgBoxResult =
        MsgBox("Apakah Anda SUDAH memiliki Log File Download ?",
               MsgBoxStyle.Question Or MsgBoxStyle.YesNo,
               "Konfirmasi Log File")

        If jawab = MsgBoxResult.No Then

            ' === BUAT FILE BARU ===
            Dim desktopPath As String =
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

            logPath = IO.Path.Combine(desktopPath, "Coretax Log Download.xlsx")

            Dim xlAppTmp As Object = CreateObject("Excel.Application")
            Dim wbTmp As Object = xlAppTmp.Workbooks.Add()
            Dim wsTmp As Object = wbTmp.Sheets(1)

            ' Header
            wsTmp.Cells(1, 1).Value = "Nomor Dokumen"
            wsTmp.Cells(1, 2).Value = "Tanggal Download"
            wsTmp.Cells(1, 3).Value = "Status"
            wsTmp.Cells(1, 4).Value = "Path File"
            wsTmp.Cells(1, 5).Value = "Nama File"

            Try
                wbTmp.SaveAs(logPath)
            Catch ex As Exception
                MsgBox("Excel kemungkinan dalam status read only, mohon di force close terlebih dahulu !")
                Exit Sub
            End Try

            wbTmp.Close(False)
            xlAppTmp.Quit()

            MsgBox("Log file berhasil dibuat di Desktop")

        Else

            ' === PILIH FILE EXISTING ===
            Dim xlTmp As Object = CreateObject("Excel.Application")
            Dim fd As Object = xlTmp.FileDialog(1)

            fd.Title = "Pilih Log File"
            fd.Filters.Clear()
            fd.Filters.Add("Excel Files", "*.xlsx")

            If fd.Show() <> -1 Then
                MsgBox("File tidak dipilih")
                xlTmp.Quit()
                Exit Sub
            End If

            logPath = fd.SelectedItems(1)
            xlTmp.Quit()

        End If

        ' ==============================
        ' BUKA EXCEL UNTUK APPEND
        ' ==============================
        Dim xlApp As Object = CreateObject("Excel.Application")
        Dim wb As Object = xlApp.Workbooks.Open(logPath)
        Dim ws As Object = wb.Sheets(1)

        Dim lastRowExcel As Integer =
        ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 1

        ' ==============================
        ' SETUP SELENIUM
        ' ==============================
        Dim driver As ChromeDriver = Nothing

        Dim dm As New DriverManager()
        dm.SetUpDriver(New ChromeConfig())

        Dim options As New ChromeOptions()

        options.AddExcludedArgument("enable-automation")
        options.AddAdditionalOption("useAutomationExtension", False)
        options.AddArgument("--disable-blink-features=AutomationControlled")
        options.AddArgument("--disable-3d-apis")
        options.AddArgument("no-sandbox")

        options.AddUserProfilePreference("profile.default_content_settings.popups", 0)
        options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1)

        options.AddUserProfilePreference("download.default_directory", "C:\DownloadPajak")
        options.AddUserProfilePreference("download.prompt_for_download", False)
        options.AddUserProfilePreference("safebrowsing.enabled", True)

        driver = New ChromeDriver(options)
        driver.Manage().Window.Maximize()

        driver.Navigate().GoToUrl("https://coretaxdjp.pajak.go.id/identityproviderportal/Account/Login/")
        Thread.Sleep(5000)

        Console.WriteLine("Step : eBupot -> Bukti Potong -> Filter -> ketik OK")

        Dim input As String = Console.ReadLine()

        If input Is Nothing OrElse input.Trim().ToUpper() <> "OK" Then
            Exit Sub
        End If

        Dim wait As New WebDriverWait(driver, TimeSpan.FromSeconds(30))

        ' ==============================
        ' VALIDASI HALAMAN
        ' ==============================
        Dim pageReady As Boolean = False

        While Not pageReady

            Console.WriteLine("Ketik OK jika sudah di halaman Export")

            Dim input2 As String = Console.ReadLine()

            If input2 IsNot Nothing AndAlso input2.Trim().ToUpper() = "OK" Then

                Try
                    Dim Pagination As IWebElement =
                    wait.Until(Function(d) d.FindElement(By.XPath("//p-paginator//p-dropdown")))

                    Pagination.Click()
                    pageReady = True

                Catch ex As Exception
                    Console.WriteLine("Export Page gak ditemukan")
                End Try

            End If

        End While

        ' ==============================
        ' SET 50 ROW
        ' ==============================
        Dim Pagination_50Pages As IWebElement =
        wait.Until(Function(d) d.FindElement(By.XPath("//p-dropdownitem[3]")))

        Pagination_50Pages.Click()

        Thread.Sleep(3000)

        ' ==============================
        ' LOOP PAGE
        ' ==============================
        Dim masihAdaPage As Boolean = True

        While masihAdaPage

            Try

                wait.Until(Function(d) d.FindElements(By.XPath("//table/tbody/tr")).Count > 0)

                Dim rows = driver.FindElements(By.XPath("//table/tbody/tr"))

                Console.WriteLine("Total Row : " & rows.Count)

                For i As Integer = 1 To rows.Count

                    Dim nomorDokumen As String = ""

                    Try
                        nomorDokumen =
                        driver.FindElement(By.XPath("//table/tbody/tr[" & i & "]/td[3]")).Text
                    Catch ex As Exception
                    End Try

                    Console.WriteLine("Download row " & i)

                    Dim statusDownload As String = "Berhasil"
                    Dim filePath As String = ""
                    Dim namaFile As String = ""

                    Try

                        Dim btnPdf As IWebElement =
                        driver.FindElement(By.XPath("//table/tbody/tr[" & i & "]/td[1]//button[2]"))

                        CType(driver, IJavaScriptExecutor).ExecuteScript("arguments[0].click();", btnPdf)

                        Thread.Sleep(3000)

                        Dim dir As New IO.DirectoryInfo("C:\DownloadPajak")
                        Dim file = dir.GetFiles().OrderByDescending(Function(f) f.LastWriteTime).FirstOrDefault()

                        If file IsNot Nothing Then
                            filePath = file.FullName
                            namaFile = IO.Path.GetFileNameWithoutExtension(file.Name)
                        End If

                    Catch ex As Exception

                        statusDownload = "Gagal"

                    End Try

                    ' ==============================
                    ' WRITE LOG
                    ' ==============================
                    ws.Cells(lastRowExcel, 1).Value = nomorDokumen
                    ws.Cells(lastRowExcel, 2).Value = Now.ToString("yyyy-MM-dd HH:mm:ss")
                    ws.Cells(lastRowExcel, 3).Value = statusDownload
                    ws.Cells(lastRowExcel, 4).Value = filePath
                    ws.Cells(lastRowExcel, 5).Value = namaFile

                    lastRowExcel += 1

                Next

            Catch ex As Exception
                Console.WriteLine("Tabel tidak ditemukan")
            End Try

            ' ==============================
            ' NEXT PAGE
            ' ==============================
            Try

                Dim btnNext As IWebElement =
                driver.FindElement(By.XPath("//p-paginator//button[2]"))

                If btnNext.Enabled Then

                    Console.WriteLine("Next Page...")
                    btnNext.Click()
                    Thread.Sleep(3000)

                Else
                    masihAdaPage = False
                End If

            Catch ex As Exception

                Console.WriteLine("Page terakhir")
                masihAdaPage = False

            End Try

        End While

        ' ==============================
        ' SAVE & CLOSE
        ' ==============================
        wb.Save()
        wb.Close(False)
        xlApp.Quit()

        ' ==============================
        ' LOGOUT
        ' ==============================
        Try
            Dim btnLogout As IWebElement =
            driver.FindElement(By.XPath("//ui-user-info//button"))

            CType(driver, IJavaScriptExecutor).ExecuteScript("arguments[0].click();", btnLogout)
        Catch ex As Exception
        End Try

        driver.Quit()

        MsgBox("Download selesai & log berhasil disimpan!")

    End Sub




    Public Sub WaitDownloadComplete(downloadFolder As String)

        Dim downloading As Boolean = True

        While downloading

            downloading = False

            Dim files = IO.Directory.GetFiles(downloadFolder, "*.crdownload")

            If files.Length > 0 Then
                downloading = True
                Threading.Thread.Sleep(1000)
            End If

        End While

    End Sub







    Sub ProsesPdfKeExcel()

        'BACA PDF TEMBAK KE EXCEL
        ' ===============================
        ' 1. SET FOLDER PDF
        ' ===============================

        Dim shell As Object = CreateObject("Shell.Application")
        Dim folder As Object = shell.BrowseForFolder(0, "Pilih folder berisi file PDF", 0)

        If folder Is Nothing Then
            MsgBox("Folder PDF tidak dipilih")
            Exit Sub
        End If

        Dim folderPath As String = folder.Self.Path

        If Not IO.Directory.Exists(folderPath) Then
            MsgBox("Folder PDF tidak ditemukan")
            Exit Sub
        End If

        Dim pdfFiles = IO.Directory.GetFiles(folderPath, "*.pdf")

        If pdfFiles.Length = 0 Then
            MsgBox("Tidak ada PDF ditemukan")
            Exit Sub
        End If

        ' ===============================
        ' 2. PILIH FILE EXCEL TUJUAN
        ' ===============================
        Dim xlApp As Object = CreateObject("Excel.Application")
        xlApp.Visible = False


        xlApp.ScreenUpdating = False
        xlApp.EnableEvents = False
        xlApp.DisplayAlerts = False

        Dim excelPath As String = ""

        Dim fd As Object = xlApp.FileDialog(1) ' Open
        fd.Filters.Clear()
        fd.Filters.Add("Excel Files", "*.xlsx;*.xls")
        fd.Title = "Pilih file Excel tujuan"

        If fd.Show() = -1 Then
            excelPath = fd.SelectedItems(1)
        Else
            MsgBox("Excel tidak dipilih")
            xlApp.Quit()
            Exit Sub
        End If

        ' ===============================
        ' 3. BUKA EXCEL
        ' ===============================
        'Dim wb As Object = xlApp.Workbooks.Open(excelPath)

        ' ===============================
        ' 3. BUKA EXCEL (AMAN)
        ' ===============================
        Dim wb As Object = Nothing

        Try
            wb = xlApp.Workbooks.Open(
        excelPath,
        ReadOnly:=False,
        IgnoreReadOnlyRecommended:=True
    )
        Catch ex As Exception
            MsgBox("Gagal membuka Excel:" & vbCrLf & ex.Message)
            xlApp.Quit()
            Exit Sub
        End Try



        ' ===============================
        ' LOAD LIST OBJEK PAJAK (1x SAJA)
        ' ===============================
        Dim wsList As Object = Nothing
        Dim dictObjek As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ' Cari sheet List Objek Pajak
        For Each sh As Object In wb.Sheets
            If sh.Name.Trim().ToUpper() = "LIST OBJEK PAJAK" Then
                wsList = sh
                Exit For
            End If
        Next

        If wsList Is Nothing Then
            MsgBox("Sheet 'List Objek Pajak' tidak ditemukan.")
            wb.Close(False)
            xlApp.Quit()
            Exit Sub
        End If

        ' Load ke Dictionary
        Dim lastRowList As Integer =
    wsList.Cells(wsList.Rows.Count, 1).End(-4162).Row ' xlUp

        For r As Integer = 2 To lastRowList
            Dim kode As String = Trim(CStr(wsList.Cells(r, 1).Value))
            Dim nama As String = Trim(CStr(wsList.Cells(r, 4).Value))

            If kode <> "" AndAlso Not dictObjek.ContainsKey(kode) Then
                dictObjek.Add(kode, nama)
            End If
        Next



        If wb Is Nothing Then
            MsgBox("Workbook tidak berhasil dibuka (wb = Nothing)")
            xlApp.Quit()
            Exit Sub
        End If


        ' ===============================
        ' 4. BUAT SHEET BARU (ddMMyyyy)
        ' ===============================
        Dim sheetName As String = DateTime.Now.ToString("ddMMyyyy")

        Dim ws As Object = Nothing

        ' Cek apakah sheet sudah ada
        For Each sh As Object In wb.Sheets
            If sh.Name = sheetName Then
                ws = sh
                Exit For
            End If
        Next



        If ws Is Nothing Then
            ' Kalau belum ada → buat baru
            ws = wb.Sheets.Add()
            ws.Name = sheetName
        Else
            ' Kalau sudah ada → HAPUS ISINYA
            ws.Cells.Clear()
        End If


        ' ===== FORMAT KOLOM TEXT (BIAR 0 TIDAK HILANG) =====
        ws.Columns(2).NumberFormat = "@"   ' NPWP WP
        ws.Columns(3).NumberFormat = "@"   ' NITKU WP
        ws.Columns(4).NumberFormat = "@"
        ws.Columns(5).NumberFormat = "@"   ' Nomor Bukti
        ws.Columns(8).NumberFormat = "@"   ' NPWP Pemotong
        ws.Columns(9).NumberFormat = "@"   ' NITKU Pemotong
        'ws.Columns(12).NumberFormat = "@"   ' DPP
        ws.Columns(13).NumberFormat = "@"   ' PPH Dipotong
        ws.Columns(14).NumberFormat = "@"   ' Tanggal Bukti Potong
        ws.Columns(15).NumberFormat = "@"
        ws.Columns(19).NumberFormat = "@"
        ws.Columns(18).NumberFormat = "@"


        ' ===== HEADER =====
        ws.Cells(1, 1).Value = "Nama Wajib Pajak (A.2)"
        ws.Cells(1, 2).Value = "NPWP / NIK WP (A.1)"
        ws.Cells(1, 3).Value = "NITKU WP (A.3)"
        ws.Cells(1, 4).Value = "Masa Pajak"
        ws.Cells(1, 5).Value = "Nomor Pemotongan"
        ws.Cells(1, 15).Value = "Status"
        ws.Cells(1, 6).Value = "Jenis Pajak"
        ws.Cells(1, 7).Value = "Kode Objek Pajak (B.4)"
        ws.Cells(1, 8).Value = "NPWP / NIK Pemotong (C.1)"
        ws.Cells(1, 9).Value = "NITKU Pemotong (C.2)"

        'ws.Cells(1, 11).Value = "Objek Pajak (B.5)"
        ws.Cells(1, 10).Value = "Nama Pemotong (C.3)"
        ws.Cells(1, 11).Value = "DPP (B.5)"
        ws.Cells(1, 12).Value = "PPh Dipotong (B.7)"
        ws.Cells(1, 14).Value = "Tanggal Bukti Potong (C.4)"
        ws.Cells(1, 13).Value = "Tarif (%) (B.6)"


        ws.Cells(1, 16).Value = "Path File PDF"

        ws.Cells(1, 17).Value = "Nama File PDF"
        ws.Cells(1, 18).Value = "Nomor Dokumen (B.9)"

        ' ===============================
        ' 5. LOOP PDF → TEMBAK KE EXCEL
        ' ===============================
        Dim row2 As Integer = 2

        For Each pdfPath As String In pdfFiles

            Dim fileName As String = IO.Path.GetFileName(pdfPath)

            ' ===== A =====
            Dim nama As String = ExtractNamaFromPdf(pdfPath)
            Dim npwp As String = ExtractNpwpFromPdf(pdfPath)
            Dim nitku As String = ExtractNitkuFromPdf(pdfPath)

            ' ===== HEADER =====
            Dim nomorMasa = ExtractNomorDanMasaPajakFromPdf(pdfPath)
            Dim nomorBukti As String = nomorMasa.Item1
            Dim masaPajak As String = nomorMasa.Item2
            Dim statusBukti As String = ExtractStatusBuktiFromPdf(pdfPath)


            ' ===== C =====
            Dim namaPemotong As String = ExtractNamaPemotongFromPdf(pdfPath)
            Dim npwpPemotong As String = ExtractNpwpPemotongFromPdf(pdfPath)
            Dim nitkuPemotong As String = ExtractNitkuPemotongFromPdf(pdfPath)
            Dim tanggalBukti As String = ExtractTanggalBuktiFromPdf(pdfPath)

            ' ===== B =====
            Dim kodeObjek As String = ExtractKodeObjekPajakFromPdf(pdfPath)
            Dim objekPajak As String = ExtractObjekPajakFromPdf(pdfPath)

            Dim dpp As String = ExtractDppFromPdf(pdfPath)
            Dim pph As String = ExtractPphFromPdf(pdfPath)
            Dim tarif As String = ExtractTarifFromPdf(pdfPath)

            Dim nomorDokumenB9 As String = ExtractNomorDokumenB9FromPdf(pdfPath)



            'Dim hasil = ExtractDppTarifPph(pdfPath)

            'Dim dpp As String = hasil.Item1
            'Dim tarif As String = hasil.Item2
            'Dim pph As String = hasil.Item3

            'Dim hasil = ExtractDppTarifPph(pdfPath)




            ' ===== WRITE EXCEL =====
            ws.Cells(row2, 1).Value = nama
            ws.Cells(row2, 2).Value = npwp
            ws.Cells(row2, 3).Value = nitku
            ws.Cells(row2, 4).Value = masaPajak
            ws.Cells(row2, 5).Value = nomorBukti

            ws.Cells(row2, 15).Value = statusBukti


            ws.Cells(row2, 7).Value = kodeObjek

            If dictObjek.ContainsKey(kodeObjek) Then
                ws.Cells(row2, 6).Value = dictObjek(kodeObjek)
            Else
                ws.Cells(row2, 6).Value = ""
            End If



            ws.Cells(row2, 8).Value = npwpPemotong
            ws.Cells(row2, 9).Value = nitkuPemotong

            ' ws.Cells(row2, 11).Value = objekPajak
            ws.Cells(row2, 10).Value = namaPemotong
            ws.Cells(row2, 11).Value = CleanNumber(dpp)
            ws.Cells(row2, 12).Value = CleanNumber(pph)
            ws.Cells(row2, 13).Value = tarif

            ws.Cells(row2, 14).Value = tanggalBukti

            ws.Cells(row2, 17).Value = fileName
            ws.Cells(row2, 16).Value = pdfPath

            ws.Cells(row2, 18).Value = nomorDokumenB9



            row2 += 1
        Next


        ' AutoFit
        ws.Columns("A:T").AutoFit()

        ' ===============================
        ' 5B. POPUP WRITE DOWN KE SHEET LAIN
        ' ===============================

        Dim targetSheetName As String =
    InputBox(
        "Tulis nama sheet untuk WRITE DOWN (append)." & vbCrLf &
        "Kosongkan jika tidak ingin copy.",
        "Write Down ke Sheet Lain",
        "Master"
    )


        If Not String.IsNullOrWhiteSpace(targetSheetName) Then

            Dim wsTarget As Object = Nothing

            ' === CARI SHEET ===
            For Each sh As Object In wb.Sheets
                If sh.Name.Trim().ToUpper() = targetSheetName.Trim().ToUpper() Then
                    wsTarget = sh
                    Exit For
                End If
            Next

            ' === JIKA BELUM ADA → BUAT ===
            If wsTarget Is Nothing Then
                wsTarget = wb.Sheets.Add()
                wsTarget.Name = targetSheetName
            End If

            ' === WAJIB: SET FORMAT SETELAH wsTarget ADA ===
            With wsTarget
                .Columns(2).NumberFormat = "@"   ' NPWP WP
                .Columns(3).NumberFormat = "@"   ' NITKU WP
                .Columns(4).NumberFormat = "@"
                .Columns(5).NumberFormat = "@"   ' Nomor Pemotongan
                .Columns(8).NumberFormat = "@"   ' NPWP Pemotong
                .Columns(9).NumberFormat = "@"   ' NITKU Pemotong
                .Columns(19).NumberFormat = "@"  ' Path
                .Columns(20).NumberFormat = "@"  ' Filename
                .Columns(18).NumberFormat = "@"  ' Filename
            End With

            If Trim(CStr(wsTarget.Cells(1, 18).Value)) <> "Nomor Dokumen" Then
                wsTarget.Cells(1, 18).Value = "Nomor Dokumen"
            End If


            ' === ROW TERAKHIR TARGET ===
            Dim lastTargetRow As Integer =
        wsTarget.Cells(wsTarget.Rows.Count, 1).End(-4162).Row
            If lastTargetRow < 2 Then lastTargetRow = 1

            ' === ROW TERAKHIR SOURCE ===
            Dim lastSourceRow As Integer =
        ws.Cells(ws.Rows.Count, 1).End(-4162).Row

            Dim nomorIndex As New Dictionary(Of String, Integer)
            Dim statusIndex As New Dictionary(Of String, String)

            For i As Integer = 2 To lastTargetRow
                Dim no As String = Trim(CStr(wsTarget.Cells(i, 5).Value))
                Dim st As String = Trim(CStr(wsTarget.Cells(i, 15).Value)).ToUpper()

                If no <> "" AndAlso Not nomorIndex.ContainsKey(no) Then
                    nomorIndex.Add(no, i)
                    statusIndex.Add(no, st) ' 🔥 INI YANG KURANG
                End If
            Next




            For r As Integer = 2 To lastSourceRow

                Dim nomorPemotongan As String = Trim(CStr(ws.Cells(r, 5).Value))
                Dim status As String = Trim(CStr(ws.Cells(r, 15).Value)).ToUpper()

                ' ===============================
                ' 1. DIBATALKAN
                ' ===============================
                If status = "DIBATALKAN" Then
                    If nomorIndex.ContainsKey(nomorPemotongan) Then
                        wsTarget.Rows(nomorIndex(nomorPemotongan)).Delete()
                        nomorIndex.Remove(nomorPemotongan)
                        statusIndex.Remove(nomorPemotongan)
                    End If
                    Continue For
                End If

                ' ===============================
                ' 2. PEMBETULAN
                ' ===============================
                'If status = "PEMBETULAN" Then

                '    ' hapus apa pun yg sudah ada
                '    If nomorIndex.ContainsKey(nomorPemotongan) Then
                '        wsTarget.Rows(nomorIndex(nomorPemotongan)).Delete()
                '        nomorIndex.Remove(nomorPemotongan)
                '        statusIndex.Remove(nomorPemotongan)
                '    End If

                '    ' insert PEMBETULAN
                '    Dim newRow As Integer =
                '        wsTarget.Cells(wsTarget.Rows.Count, 1).End(-4162).Row + 1

                '    For c As Integer = 1 To 20
                '        wsTarget.Cells(newRow, c).Value = ws.Cells(r, c).Value
                '    Next

                '    nomorIndex.Add(nomorPemotongan, newRow)
                '    statusIndex.Add(nomorPemotongan, "PEMBETULAN")

                '    Continue For
                'End If

                If status = "PEMBETULAN" Then

                    Dim sourceData As String = JoinRow(ws, r, 20)
                    Dim isDuplicate As Boolean = False

                    ' Ambil last row master (DECLARE SEKALI SAJA)
                    Dim lastMasterRow As Integer =
        wsTarget.Cells(wsTarget.Rows.Count, 5).End(-4162).Row

                    ' =========================
                    ' 1. CEK DUPLICATE DULU
                    ' =========================
                    For i As Integer = 2 To lastMasterRow

                        Dim nomorMaster As String =
            Trim(CStr(wsTarget.Cells(i, 5).Value2))

                        If nomorMaster = nomorPemotongan Then

                            Dim masterData As String =
                JoinRow(wsTarget, i, 20)

                            If sourceData = masterData Then
                                isDuplicate = True
                                Exit For
                            End If

                        End If

                    Next

                    ' Kalau duplicate → STOP
                    If isDuplicate Then
                        Continue For
                    End If


                    ' =========================
                    ' 2. HAPUS SEMUA NOMOR SAMA
                    ' =========================
                    For i As Integer = lastMasterRow To 2 Step -1

                        Dim nomorMaster As String =
            Trim(CStr(wsTarget.Cells(i, 5).Value2))

                        Dim statusMaster As String =
                            Trim(CStr(wsTarget.Cells(i, 15).Value2)).ToUpper()

                        If nomorMaster = nomorPemotongan AndAlso statusMaster = "NORMAL" Then
                            wsTarget.Rows(i).Delete()
                        End If


                    Next


                    ' =========================
                    ' 3. INSERT PEMBETULAN
                    ' =========================
                    Dim newRow As Integer =
        wsTarget.Cells(wsTarget.Rows.Count, 1).End(-4162).Row + 1

                    For c As Integer = 1 To 20
                        wsTarget.Cells(newRow, c).Value =
            ws.Cells(r, c).Value
                    Next

                    Continue For

                End If



                ' INSERT

                ' ===============================
                ' 3. NORMAL
                ' ===============================
                If status = "NORMAL" Then

                    ' sudah ada PEMBETULAN → skip
                    If statusIndex.ContainsKey(nomorPemotongan) _
                       AndAlso statusIndex(nomorPemotongan) = "PEMBETULAN" Then
                        Continue For
                    End If

                    ' sudah ada data → skip
                    If nomorIndex.ContainsKey(nomorPemotongan) Then
                        Continue For
                    End If

                    ' insert NORMAL
                    Dim newRow As Integer =
                        wsTarget.Cells(wsTarget.Rows.Count, 1).End(-4162).Row + 1

                    For c As Integer = 1 To 20
                        wsTarget.Cells(newRow, c).Value = ws.Cells(r, c).Value
                    Next

                    nomorIndex.Add(nomorPemotongan, newRow)
                    statusIndex.Add(nomorPemotongan, "NORMAL")
                End If

            Next


        End If


        ' ===============================
        ' 6. RESTORE EXCEL STATE
        ' ===============================
        If xlApp IsNot Nothing Then
            xlApp.ScreenUpdating = True
            xlApp.EnableEvents = True
            xlApp.DisplayAlerts = True
        End If

        ' ===============================
        ' 7. SAVE & CLEANUP
        ' ===============================
        wb.Save()
        wb.Close(False)
        xlApp.Quit()

        Marshal.ReleaseComObject(ws)
        Marshal.ReleaseComObject(wb)
        Marshal.ReleaseComObject(xlApp)

        ws = Nothing
        wb = Nothing
        xlApp = Nothing

        GC.Collect()
        GC.WaitForPendingFinalizers()

        MsgBox("Selesai ✔ PDF berhasil ditembak ke Excel")


    End Sub



    Public Sub HighlightElement(driver As IWebDriver, element As IWebElement)
        Dim js As IJavaScriptExecutor = CType(driver, IJavaScriptExecutor)
        js.ExecuteScript(
            "arguments[0].style.border='3px solid red';" &
            "arguments[0].style.backgroundColor='#fff3cd';",
            element)
    End Sub

    Function ExtractNamaFromPdf(pdfPath As String) As String

        Dim fullText As New Text.StringBuilder()

        Using reader As New PdfReader(pdfPath)
            Using pdf As New PdfDocument(reader)

                For i As Integer = 1 To pdf.GetNumberOfPages()
                    fullText.Append(
                        PdfTextExtractor.GetTextFromPage(pdf.GetPage(i))
                    )
                Next

            End Using
        End Using

        Dim text As String = fullText.ToString()

        Dim match As Match =
            Regex.Match(text, "A\.2\s+NAMA\s*:\s*(.+)")

        If match.Success Then
            Return match.Groups(1).Value.Trim()
        End If

        Return ""
    End Function

    Function ExtractNpwpFromPdf(pdfPath As String) As String

        Dim fullText As New Text.StringBuilder()

        Using reader As New PdfReader(pdfPath)
            Using pdf As New PdfDocument(reader)
                For i As Integer = 1 To pdf.GetNumberOfPages()
                    fullText.Append(
                    PdfTextExtractor.GetTextFromPage(pdf.GetPage(i))
                )
                Next
            End Using
        End Using

        Dim text As String = fullText.ToString()

        ' Ambil angka setelah "A.1 NPWP / NIK :"
        Dim match As Match =
        Regex.Match(
            text,
            "A\.1\s+NPWP\s*/\s*NIK\s*:\s*([0-9]+)"
        )

        If match.Success Then
            Return match.Groups(1).Value.Trim()
        End If

        Return ""
    End Function

    Function ExtractNitkuFromPdf(pdfPath As String) As String

        Dim fullText As New Text.StringBuilder()

        Using reader As New PdfReader(pdfPath)
            Using pdf As New PdfDocument(reader)
                For i As Integer = 1 To pdf.GetNumberOfPages()
                    fullText.Append(
                        PdfTextExtractor.GetTextFromPage(pdf.GetPage(i))
                    )
                Next
            End Using
        End Using

        Dim text As String = fullText.ToString()

        ' STRATEGI: cari "A.3", lalu ":" lalu AMBIL ANGKA SEBELUM "-"
        Dim match As Match =
            Regex.Match(
                text,
                "A\.?3.*?:\s*([0-9]{10,})",
                RegexOptions.IgnoreCase Or RegexOptions.Singleline
            )

        If match.Success Then
            Return match.Groups(1).Value.Trim()
        End If

        Return ""
    End Function

    Function ExtractMasaPajakFromPdf(pdfPath As String) As String

        Dim fullText As New Text.StringBuilder()

        Using reader As New PdfReader(pdfPath)
            Using pdf As New PdfDocument(reader)
                For i As Integer = 1 To pdf.GetNumberOfPages()
                    fullText.AppendLine(
                    PdfTextExtractor.GetTextFromPage(pdf.GetPage(i))
                )
                Next
            End Using
        End Using

        Dim text As String = fullText.ToString()

        ' Ambil baris yang berisi NOMOR MASA PAJAK
        Dim match As Match =
        Regex.Match(
            text,
            "NOMOR\s+MASA\s+PAJAK[\s\S]*?\n\s*([A-Z0-9]+)\s+(\d{2}\-\d{4})",
            RegexOptions.IgnoreCase
        )

        If match.Success Then
            ' Group 2 = MM-YYYY
            Return match.Groups(2).Value.Trim()
        End If

        Return ""
    End Function

    Function ExtractNomorDanMasaPajakFromPdf(pdfPath As String) As Tuple(Of String, String)

        Dim fullText As New Text.StringBuilder()

        Using reader As New PdfReader(pdfPath)
            Using pdf As New PdfDocument(reader)
                For i As Integer = 1 To pdf.GetNumberOfPages()
                    fullText.AppendLine(
                        PdfTextExtractor.GetTextFromPage(pdf.GetPage(i))
                    )
                Next
            End Using
        End Using

        Dim text As String = fullText.ToString()

        Dim match As Match =
            Regex.Match(
                text,
                "NOMOR\s+MASA\s+PAJAK[\s\S]*?\n\s*([A-Z0-9]+)\s+(\d{2}\-\d{4})",
                RegexOptions.IgnoreCase
            )

        If match.Success Then
            Dim nomorBukti As String = match.Groups(1).Value.Trim()
            Dim masaPajak As String = match.Groups(2).Value.Trim()

            Return Tuple.Create(nomorBukti, masaPajak)
        End If

        Return Tuple.Create("", "")
    End Function

    Function ExtractKodeObjekPajakFromPdf(pdfPath As String) As String

        Dim fullText As New Text.StringBuilder()

        Using reader As New PdfReader(pdfPath)
            Using pdf As New PdfDocument(reader)
                For i As Integer = 1 To pdf.GetNumberOfPages()
                    fullText.AppendLine(
                    PdfTextExtractor.GetTextFromPage(pdf.GetPage(i))
                )
                Next
            End Using
        End Using

        Dim text As String = fullText.ToString()

        ' Cari pola KODE OBJEK PAJAK: NN-NNN-NN
        Dim match As Match =
        Regex.Match(
            text,
            "\b(\d{2}\-\d{3}\-\d{2})\b"
        )

        If match.Success Then
            Return match.Groups(1).Value.Trim()
        End If

        Return ""
    End Function

    Function ExtractNpwpPemotongFromPdf(pdfPath As String) As String

        Dim fullText As New Text.StringBuilder()

        Using reader As New PdfReader(pdfPath)
            Using pdf As New PdfDocument(reader)
                For i As Integer = 1 To pdf.GetNumberOfPages()
                    fullText.AppendLine(
                    PdfTextExtractor.GetTextFromPage(pdf.GetPage(i))
                )
                Next
            End Using
        End Using

        Dim text As String = fullText.ToString()

        ' Ambil NPWP/NIK pada C.1
        Dim match As Match =
        Regex.Match(
            text,
            "C\.1\s+NPWP\s*/\s*NIK\s*:\s*([0-9]+)",
            RegexOptions.IgnoreCase
        )

        If match.Success Then
            Return match.Groups(1).Value.Trim()
        End If

        Return ""
    End Function

    Function ExtractNitkuPemotongFromPdf(pdfPath As String) As String

        Dim fullText As New Text.StringBuilder()

        Using reader As New PdfReader(pdfPath)
            Using pdf As New PdfDocument(reader)
                For i As Integer = 1 To pdf.GetNumberOfPages()
                    fullText.AppendLine(
                    PdfTextExtractor.GetTextFromPage(pdf.GetPage(i))
                )
                Next
            End Using
        End Using

        Dim text As String = fullText.ToString()

        Dim match As Match =
        Regex.Match(
            text,
            "C\.2[\s\S]*?:\s*([0-9]{10,})",
            RegexOptions.IgnoreCase Or RegexOptions.Singleline
        )

        If match.Success Then
            Return match.Groups(1).Value.Trim()
        End If

        Return ""
    End Function

    Function ExtractNamaPemotongFromPdf(pdfPath As String) As String

        Dim text As String = ReadPdfText(pdfPath)

        Dim match As Match =
        Regex.Match(
            text,
            "C\.3\s+NAMA\s+PEMOTONG[\s\S]*?:\s*(.+)",
            RegexOptions.IgnoreCase
        )

        If match.Success Then
            Return match.Groups(1).Value.Trim()
        End If

        Return ""
    End Function

    Function ExtractObjekPajakFromPdf(pdfPath As String) As String

        Dim text As String = ReadPdfText(pdfPath)

        Dim match As Match =
        Regex.Match(
            text,
            "\d{2}-\d{3}-\d{2}\s+(.+?)\s+\d{1,3}(\.\d{3})+",
            RegexOptions.IgnoreCase
        )

        If match.Success Then
            Return match.Groups(1).Value.Trim()
        End If

        Return ""
    End Function



    Function ExtractDppFromPdf(pdfPath As String) As String

        Dim text As String = ReadPdfText(pdfPath)

        ' cari kode objek pajak dulu
        Dim mKode As Match =
        Regex.Match(text, "\b\d{2}-\d{3}-\d{2}\b")

        If Not mKode.Success Then Return ""

        ' ambil text SETELAH kode objek pajak
        Dim afterKode As String =
        text.Substring(mKode.Index + mKode.Length)

        ' ambil semua angka format rupiah
        Dim angkaMatches As MatchCollection =
        Regex.Matches(afterKode, "\b\d{1,3}(?:\.\d{3})+\b")

        ' minimal harus ada DPP
        If angkaMatches.Count = 0 Then Return ""

        ' index 0 = DPP
        Return angkaMatches(0).Value.Trim()

    End Function





    'Function ExtractDppFromPdf(pdfPath As String) As String


    '    Dim text As String = ReadPdfText(pdfPath)

    '    Dim match As Match =
    '    Regex.Match(
    '        text,
    '        "\d{2}-\d{3}-\d{2}[\s\S]+?(\d{1,3}(?:\.\d{3})*(?:,\d+)?)",
    '        RegexOptions.IgnoreCase
    '    )

    '    If match.Success Then
    '        Return match.Groups(1).Value.Trim()
    '    End If

    '    Return ""
    'End Function



    'Function ExtractDppTarifPph(pdfPath As String) As Tuple(Of String, String, String)

    '    Dim text As String = ReadPdfText(pdfPath)

    '    Dim lines() As String =
    '    text.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

    '    For i As Integer = 0 To lines.Length - 1

    '        If Regex.IsMatch(lines(i), "\b\d{2}-\d{3}-\d{2}\b") Then

    '            Dim angka As New List(Of String)

    '            For j As Integer = i + 1 To Math.Min(i + 10, lines.Length - 1)

    '                Dim l As String = lines(j).Trim()

    '                If Regex.IsMatch(l, "^\d{1,3}(?:\.\d{3})*(?:,\d+)?$") Then
    '                    angka.Add(l)
    '                End If

    '                If angka.Count = 3 Then

    '                    Dim dpp As String = ""
    '                    Dim tarif As String = ""
    '                    Dim pph As String = ""

    '                    For Each a In angka
    '                        Dim clean = a.Replace(".", "").Replace(",", ".")
    '                        Dim val As Decimal

    '                        If Decimal.TryParse(clean, val) Then
    '                            If val <= 100 Then
    '                                tarif = a
    '                            ElseIf dpp = "" Then
    '                                dpp = a
    '                            Else
    '                                pph = a
    '                            End If
    '                        End If
    '                    Next

    '                    Return Tuple.Create(dpp, tarif, pph)
    '                End If
    '            Next
    '        End If
    '    Next

    '    Return Tuple.Create("", "", "")
    'End Function














    Function extractdpptarifpph(pdfpath As String) As Tuple(Of String, String, String)

        Dim text As String = ReadPdfText(pdfpath)

        Dim lines() As String =
        text.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

        For i As Integer = 0 To lines.Length - 1

            ' ketemu kode objek pajak
            If Regex.IsMatch(lines(i), "\b\d{2}-\d{3}-\d{2}\b") Then

                Dim angka As New List(Of String)

                For j As Integer = i + 1 To Math.Min(i + 10, lines.Length - 1)

                    Dim l As String = lines(j).Trim()

                    ' === case 1: 1 baris 3 angka (pasal 22) ===
                    Dim inlinematch As Match =
                    Regex.Match(
                        l,
                        "^(\d{1,3}(?:\.\d{3})+)\s+([\d.,]+)\s+(\d{1,3}(?:\.\d{3})+)$"
                    )

                    If inlinematch.Success Then
                        Return Tuple.Create(
                        inlinematch.Groups(1).Value,
                        inlinematch.Groups(2).Value,
                        inlinematch.Groups(3).Value
                    )
                    End If

                    ' === case 2: baris angka terpisah (pasal 23) ===
                    If Regex.IsMatch(l, "^\d{1,3}(?:\.\d{3})*(?:,\d+)?$") Then
                        angka.Add(l)
                    End If

                    If angka.Count = 3 Then
                        Return Tuple.Create(angka(0), angka(1), angka(2))
                    End If

                Next
            End If
        Next

        Return Tuple.Create("", "", "")
    End Function

    Function ExtractPphFromPdf(pdfPath As String) As String

        Dim text As String = ReadPdfText(pdfPath)

        ' 1. Cari kode objek pajak
        Dim mKode As Match =
        Regex.Match(text, "\b\d{2}-\d{3}-\d{2}\b")

        If Not mKode.Success Then Return ""

        ' 2. Ambil text setelah kode objek pajak
        Dim afterKode As String =
        text.Substring(mKode.Index + mKode.Length)

        ' 3. Ambil semua angka
        Dim angkaMatches As MatchCollection =
        Regex.Matches(
            afterKode,
            "\b\d{1,3}(?:\.\d{3})+\b|\b\d+(?:[.,]\d+)?\b"
        )

        ' 4. Scan berurutan: DPP - Tarif - PPh
        For i As Integer = 0 To angkaMatches.Count - 3

            Dim dppVal As String = angkaMatches(i).Value
            Dim tarifVal As String = angkaMatches(i + 1).Value
            Dim pphVal As String = angkaMatches(i + 2).Value

            ' DPP wajib ribuan
            If Not dppVal.Contains(".") Then Continue For

            ' Tarif harus kecil & masuk akal
            Dim tarif As Decimal
            If Not Decimal.TryParse(tarifVal.Replace(",", "."), tarif) Then Continue For
            If tarif <= 0 OrElse tarif > 100 Then Continue For

            ' === INI KUNCINYA ===
            ' PPh BOLEH kecil / tanpa titik
            Return pphVal.Replace(".", "").Replace(",", "").Trim()

        Next

        Return ""

    End Function










    'Function ExtractPphFromPdf(pdfPath As String) As String

    '    Dim text As String = ReadPdfText(pdfPath)

    '    ' 1. Cari kode objek pajak (B.4)
    '    Dim mKode As Match =
    '    Regex.Match(text, "\b\d{2}-\d{3}-\d{2}\b")

    '    If Not mKode.Success Then Return ""

    '    ' 2. Ambil text SETELAH kode objek pajak
    '    Dim afterKode As String =
    '    text.Substring(mKode.Index + mKode.Length)

    '    ' 3. Ambil angka berurutan (DPP | Tarif | PPh)
    '    Dim angkaMatches As MatchCollection =
    '    Regex.Matches(
    '        afterKode,
    '        "\b\d{1,3}(?:\.\d{3})+\b|\b\d+(?:\.\d+)?\b"
    '    )

    '    ' 4. Minimal harus ada 3 angka
    '    If angkaMatches.Count < 3 Then Return ""

    '    ' Urutan:
    '    ' index 0 = DPP
    '    ' index 1 = Tarif
    '    ' index 2 = PPh
    '    Dim pphRaw As String = angkaMatches(2).Value

    '    ' 5. Normalisasi angka
    '    pphRaw = pphRaw.Replace(".", "").Replace(",", "").Trim()

    '    Return pphRaw

    'End Function




    Function ExtractTanggalBuktiFromPdf(pdfPath As String) As String

        Dim text As String = ReadPdfText(pdfPath)

        Dim match As Match =
        Regex.Match(
            text,
            "C\.4\s+TANGGAL\s*:\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
            RegexOptions.IgnoreCase
        )

        If match.Success Then
            Return match.Groups(1).Value.Trim()
        End If

        Return ""
    End Function

    'Function ExtractStatusBuktiFromPdf(pdfPath As String) As String

    '    Dim text As String = ReadPdfText(pdfPath)

    '    ' Cari kata setelah "STATUS BUKTI PEMOTONGAN"
    '    Dim match As Match =
    '    Regex.Match(
    '        text,
    '        "STATUS\s+BUKTI\s+PEMOTONGAN[\s\S]{0,100}?\b(NORMAL|DIBATALKAN)\b",
    '        RegexOptions.IgnoreCase
    '    )

    '    If match.Success Then
    '        Return match.Groups(1).Value.Trim().ToUpper()
    '    End If

    '    Return ""
    'End Function

    Function ExtractStatusBuktiFromPdf(pdfPath As String) As String

        Dim text As String = ReadPdfText(pdfPath)

        ' Cari status di baris NOMOR MASA PAJAK
        Dim match As Match =
        Regex.Match(
            text,
            "\b(NORMAL|DIBATALKAN|PEMBETULAN)\b",
            RegexOptions.IgnoreCase
        )

        If match.Success Then
            Return match.Groups(1).Value.Trim().ToUpper()
        End If

        Return ""

    End Function




    Function ReadPdfText(pdfPath As String) As String

        Dim sb As New Text.StringBuilder()

        Using reader As New PdfReader(pdfPath)
            Using pdf As New PdfDocument(reader)
                For i As Integer = 1 To pdf.GetNumberOfPages()
                    sb.AppendLine(
                    PdfTextExtractor.GetTextFromPage(pdf.GetPage(i))
                )
                Next
            End Using
        End Using

        Return sb.ToString()
    End Function


    Function ExtractTarifFromPdf(pdfPath As String) As String

        Dim text As String = ReadPdfText(pdfPath)

        ' 1. Cari kode objek pajak
        Dim mKode As Match =
        Regex.Match(text, "\b\d{2}-\d{3}-\d{2}\b")

        If Not mKode.Success Then Return ""

        ' 2. Ambil text setelah kode objek pajak
        Dim afterKode As String =
        text.Substring(mKode.Index + mKode.Length)

        ' 3. Ambil semua angka (ribuan & desimal titik/koma)
        Dim angkaMatches As MatchCollection =
        Regex.Matches(
            afterKode,
            "\b\d{1,3}(?:\.\d{3})+\b|\b\d+(?:[.,]\d+)?\b"
        )

        ' 4. Cari pola: DPP (besar) - TARIF (kecil) - PPH (bebas)
        For i As Integer = 0 To angkaMatches.Count - 3

            Dim dppVal As String = angkaMatches(i).Value
            Dim tarifVal As String = angkaMatches(i + 1).Value

            ' DPP wajib ribuan
            If Not dppVal.Contains(".") Then Continue For

            ' Parse tarif (support 0.1 dan 0,1)
            Dim tarif As Decimal
            If Not Decimal.TryParse(tarifVal.Replace(",", "."), tarif) Then Continue For

            ' Tarif masuk akal
            If tarif > 0 AndAlso tarif <= 100 Then
                Return tarifVal.Trim()
            End If

        Next

        Return ""

    End Function






    'Function ExtractTarifFromPdf(pdfPath As String) As String

    '    Dim text As String = ReadPdfText(pdfPath)

    '    ' 1. Cari posisi kode objek pajak
    '    Dim mKode As Match =
    '    Regex.Match(text, "\d{2}-\d{3}-\d{2}")

    '    If Not mKode.Success Then Return ""

    '    ' 2. Ambil text setelah kode objek pajak
    '    Dim afterKode As String =
    '    text.Substring(mKode.Index + mKode.Length)

    '    ' 3. Ambil SEMUA angka (DPP / Tarif / PPh)
    '    Dim angkaMatches As MatchCollection =
    '    Regex.Matches(
    '        afterKode,
    '        "\b\d{1,3}(?:\.\d{3})+\b|\b\d+(\.\d+)?\b"
    '    )

    '    ' 4. Pastikan minimal 3 angka
    '    If angkaMatches.Count >= 3 Then
    '        ' urutan: DPP | TARIF | PPH
    '        Return angkaMatches(1).Value.Trim()
    '    End If

    '    Return ""
    'End Function


    Function CleanNumber(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return ""

        Return text.Replace(".", "").Replace(",", "").Trim()
    End Function

    Function ExistsNomorPemotongan(wsTarget As Object, nomorPemotongan As String) As Boolean
        Dim lastRow As Integer =
            wsTarget.Cells(wsTarget.Rows.Count, 5).End(-4162).Row ' xlUp

        For r As Integer = 2 To lastRow
            If wsTarget.Cells(r, 5).Value IsNot Nothing Then
                If wsTarget.Cells(r, 5).Value.ToString().Trim() = nomorPemotongan Then
                    Return True
                End If
            End If
        Next

        Return False
    End Function

    Sub DeleteByNomorPemotongan(wsTarget As Object, nomorPemotongan As String)
        Dim lastRow As Integer =
        wsTarget.Cells(wsTarget.Rows.Count, 5).End(-4162).Row ' xlUp

        For r As Integer = lastRow To 2 Step -1
            If wsTarget.Cells(r, 5).Value IsNot Nothing Then
                If wsTarget.Cells(r, 5).Value.ToString().Trim() = nomorPemotongan Then
                    wsTarget.Rows(r).Delete()
                End If
            End If
        Next
    End Sub



    Function IsFileLocked(path As String) As Boolean
        Try
            Using fs = IO.File.Open(path, IO.FileMode.Open, IO.FileAccess.ReadWrite, IO.FileShare.None)
            End Using
            Return False
        Catch
            Return True
        End Try
    End Function



    'Function GetLatestPdf(downloadDir As String, beforeFiles As HashSet(Of String)) As String
    '    Dim files = Directory.GetFiles(downloadDir, "*.pdf")

    '    For Each f In files
    '        If Not beforeFiles.Contains(f) Then
    '            Return f
    '        End If
    '    Next

    '    Return ""
    'End Function


    Function GetLatestPdf(downloadDir As String, beforeFiles As HashSet(Of String)) As String

        Dim files = System.IO.Directory.GetFiles(downloadDir, "*.pdf") _
        .OrderByDescending(Function(f) System.IO.File.GetCreationTime(f))

        For Each f As String In files
            If Not beforeFiles.Contains(f) Then
                Return f
            End If
        Next

        Return ""

    End Function






    'Function ExtractNomorDokumenB9FromPdf(pdfPath As String) As String

    '    Dim text As String = ReadPdfText(pdfPath)

    '    Dim match As Match =
    '    Regex.Match(
    '        text,
    '        "B\.9\s+Nomor\s+Dokumen\s*:\s*([A-Z0-9\-]+)",
    '        RegexOptions.IgnoreCase
    '    )

    '    If match.Success Then
    '        Return match.Groups(1).Value.Trim()
    '    End If

    '    Return ""

    'End Function



    Function ExtractNomorDokumenB9FromPdf(pdfPath As String) As String

        Dim text As String = ReadPdfText(pdfPath)

        Dim match As Match =
        Regex.Match(
            text,
            "B\.9\s+Nomor\s+Dokumen\s*:\s*([^\r\n]+)",
            RegexOptions.IgnoreCase
        )

        If match.Success Then
            Return match.Groups(1).Value.Trim()
        End If

        Return ""

    End Function



    Function JoinRow(ws As Excel.Worksheet, rowNum As Integer, maxCol As Integer) As String

        Dim arr As New List(Of String)

        For c As Integer = 1 To maxCol

            ' 🚫 Skip kolom P (16) dan Q (17)
            If c = 16 Or c = 17 Then
                Continue For
            End If

            Dim val = ws.Cells(rowNum, c).Value2

            If val Is Nothing Then
                arr.Add("")
            ElseIf IsNumeric(val) Then
                arr.Add(Convert.ToDecimal(val).ToString("0.################"))
            Else
                arr.Add(Trim(CStr(val)))
            End If

        Next

        Return String.Join("|", arr)

    End Function









End Module