'20151025 - ProcessTimeouts Program
'Use to move timeouts from the backup directories to their main directories to be reprocessed.
'Run this program periodically from the system scheduler.

'20151216 - add INVULHMCARE processing.
'20160412 - removed IVNULHMCARE due to switch to STAR.
'201404121257 - added processing for ULH and ULHORU back in - Gary

Imports System
Imports System.IO
Imports System.Collections
Imports System.Data.SqlClient

Module Module1
    Public objIniFile As New INIFile("d:\W3Production\HL7Mapper.ini")
    Dim strLogDirectory As String = ""
    Dim functionError As Boolean = False
    Dim gblLogString As String = ""

    Sub main()
        Try
            Call processMcare()
            Call processMcareORU()
            Call processMEDITECH()
            Call processShelby()
            Call processShelbyORU()
            Call processITW()
            Call processITWORM()
            Call processITWORU()
            Call processITWSIU()
            Call processULH() '20160412
            Call processULHORU() '20160412

            '20160412 - removed IVNULHMCARE due to switch to STAR.
            'Call processINVULHMCARE() '20151216 -  added

        Finally
            If functionError Then
                writeTolog(gblLogString, 1)
            End If
        End Try
    End Sub

    Sub processINVULHMCARE()
        Try
            Dim dir As String
            Dim thefile

            strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)")
            Dim strOutputDirectory As String = objIniFile.GetString("INVULHMCARE", "INVULHMCAREoutputdirectory", "(none)") 'Like d:\W3Feeds\NVP\Mcare\
            Dim strBackupDirectory As String = strOutputDirectory & "backup\"
            Dim strReprocessDirectory As String = strOutputDirectory & "reprocess\"

            'move the files in backup to the output directory
            Dim dirs As String() = Directory.GetFiles(strBackupDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

            'move files in reprocess to output directory
            dirs = Directory.GetFiles(strReprocessDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "INVULHMCARE Process Timeout Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

        Finally

        End Try


    End Sub

    Sub processULH()
        Try
            Dim dir As String
            Dim thefile

            strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)")
            Dim strOutputDirectory As String = objIniFile.GetString("ULH", "ULHoutputdirectory", "(none)") 'Like d:\W3Feeds\NVP\Mcare\
            Dim strBackupDirectory As String = strOutputDirectory & "backup\"
            Dim strReprocessDirectory As String = strOutputDirectory & "reprocess\"

            'move the files in backup to the output directory
            Dim dirs As String() = Directory.GetFiles(strBackupDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

            'move files in reprocess to output directory
            dirs = Directory.GetFiles(strReprocessDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "ULH Process Timeout Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

        Finally

        End Try


    End Sub

    Sub processULHORU()
        Try
            Dim dir As String
            Dim thefile

            strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)")
            Dim strOutputDirectory As String = objIniFile.GetString("ULHORU", "ULHORUoutputdirectory", "(none)") 'Like d:\W3Feeds\NVP\Mcare\
            Dim strBackupDirectory As String = strOutputDirectory & "backup\"
            Dim strReprocessDirectory As String = strOutputDirectory & "reprocess\"

            'move the files in backup to the output directory
            Dim dirs As String() = Directory.GetFiles(strBackupDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

            'move files in reprocess to output directory
            dirs = Directory.GetFiles(strReprocessDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "ULHORU Process Timeout Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

        Finally

        End Try


    End Sub

    Sub processITW()
        Try
            Dim dir As String
            Dim thefile

            strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)")
            Dim strOutputDirectory As String = objIniFile.GetString("ITW", "ITWoutputdirectory", "(none)") 'Like d:\W3Feeds\NVP\Mcare\
            Dim strBackupDirectory As String = strOutputDirectory & "backup\"
            Dim strReprocessDirectory As String = strOutputDirectory & "reprocess\"
            '6/27/2017 - add orphans - ULHT
            Dim strorphanDirectory As String = strOutputDirectory & "orphans\"

            'move the files in backup to the output directory
            Dim dirs As String() = Directory.GetFiles(strBackupDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

            'move files in reprocess to output directory
            dirs = Directory.GetFiles(strReprocessDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

            '6/27/2017 - add orphans - ULHT
            'move files in orphans to output directory
            dirs = Directory.GetFiles(strorphanDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next



        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "ITW Process Timeout Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

        Finally

        End Try


    End Sub

    Sub processITWORM()
        Try
            Dim dir As String
            Dim thefile

            strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)")
            Dim strOutputDirectory As String = objIniFile.GetString("ITWORM", "ITWORMoutputdirectory", "(none)") 'Like d:\W3Feeds\NVP\Mcare\
            Dim strBackupDirectory As String = strOutputDirectory & "backup\"
            Dim strReprocessDirectory As String = strOutputDirectory & "reprocess\"

            'move the files in backup to the output directory
            Dim dirs As String() = Directory.GetFiles(strBackupDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

            'move files in reprocess to output directory
            dirs = Directory.GetFiles(strReprocessDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "ITWORM Process Timeout Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

        Finally

        End Try


    End Sub

    Sub processITWORU()
        Try
            Dim dir As String
            Dim thefile

            strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)")
            Dim strOutputDirectory As String = objIniFile.GetString("ITWORU", "ITWORUoutputdirectory", "(none)") 'Like d:\W3Feeds\NVP\Mcare\
            Dim strBackupDirectory As String = strOutputDirectory & "backup\"
            Dim strReprocessDirectory As String = strOutputDirectory & "reprocess\"

            'move the files in backup to the output directory
            Dim dirs As String() = Directory.GetFiles(strBackupDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

            'move files in reprocess to output directory
            dirs = Directory.GetFiles(strReprocessDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "ITWORU Process Timeout Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

        Finally

        End Try


    End Sub

    Sub processITWSIU()
        Try
            Dim dir As String
            Dim thefile

            strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)")
            Dim strOutputDirectory As String = objIniFile.GetString("ITWSIU", "ITWSIUoutputdirectory", "(none)") 'Like d:\W3Feeds\NVP\Mcare\
            Dim strBackupDirectory As String = strOutputDirectory & "backup\"
            Dim strReprocessDirectory As String = strOutputDirectory & "reprocess\"

            'move the files in backup to the output directory
            Dim dirs As String() = Directory.GetFiles(strBackupDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

            'move files in reprocess to output directory
            dirs = Directory.GetFiles(strReprocessDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "ITWSIU Process Timeout Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

        Finally

        End Try


    End Sub

    Sub processMcare()
        Try
            Dim dir As String
            Dim thefile

            strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)")
            Dim strOutputDirectory As String = objIniFile.GetString("MCARE", "Mcareoutputdirectory", "(none)") 'Like d:\W3Feeds\NVP\Mcare\
            Dim strBackupDirectory As String = strOutputDirectory & "backup\"
            Dim strReprocessDirectory As String = strOutputDirectory & "reprocess\"

            'move the files in backup to the output directory
            Dim dirs As String() = Directory.GetFiles(strBackupDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

            'move files in reprocess to output directory
            dirs = Directory.GetFiles(strReprocessDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Mcare Process Timeout Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

        Finally

        End Try


    End Sub

    Sub processShelby()
        Try
            Dim dir As String
            Dim thefile

            strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)")
            Dim strOutputDirectory As String = objIniFile.GetString("Shelby", "Shelbyoutputdirectory", "(none)") 'Like d:\W3Feeds\NVP\Mcare\
            Dim strBackupDirectory As String = strOutputDirectory & "backup\"
            Dim strReprocessDirectory As String = strOutputDirectory & "reprocess\"

            'move the files in backup to the output directory
            Dim dirs As String() = Directory.GetFiles(strBackupDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

            'move files in reprocess to output directory
            dirs = Directory.GetFiles(strReprocessDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Shelby Process Timeout Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

        Finally

        End Try


    End Sub
    Sub processMcareORU()
        Try
            Dim dir As String
            Dim thefile

            strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)")
            Dim strOutputDirectory As String = objIniFile.GetString("MCAREORU", "MCAREORUoutputdirectory", "(none)") 'Like d:\W3Feeds\NVP\Mcare\
            Dim strBackupDirectory As String = strOutputDirectory & "backup\"
            Dim strReprocessDirectory As String = strOutputDirectory & "reprocess\"

            'move the files in backup to the output directory
            Dim dirs As String() = Directory.GetFiles(strBackupDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

            'move files in reprocess to output directory
            dirs = Directory.GetFiles(strReprocessDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "McareORU Process Timeout Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

        Finally

        End Try


    End Sub

    Sub processShelbyORU()
        Try
            Dim dir As String
            Dim thefile

            strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)")
            Dim strOutputDirectory As String = objIniFile.GetString("SHELBYORU", "SHELBYORUoutputdirectory", "(none)") 'Like d:\W3Feeds\NVP\Mcare\
            Dim strBackupDirectory As String = strOutputDirectory & "backup\"
            Dim strReprocessDirectory As String = strOutputDirectory & "reprocess\"

            'move the files in backup to the output directory
            Dim dirs As String() = Directory.GetFiles(strBackupDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

            'move files in reprocess to output directory
            dirs = Directory.GetFiles(strReprocessDirectory, "NVP.*")
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "ShelbyORU Process Timeout Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

        Finally

        End Try


    End Sub

    Sub processMEDITECH()
        Try
            Dim dir As String
            Dim thefile

            strLogDirectory = objIniFile.GetString("Settings", "logs", "(none)")
            Dim strOutputDirectory As String = objIniFile.GetString("CLOVERLEAF", "CLOVERLEAFoutputdirectory", "(none)") 'Like d:\W3Feeds\NVP\Mcare\
            Dim strBackupDirectory As String = strOutputDirectory & "backup\"
            Dim strReprocessDirectory As String = strOutputDirectory & "reprocess\"

            'move the files in backup to the output directory
            Dim dirs As String() = Directory.GetFiles(strBackupDirectory, "LTW.*") 'Note: Cloverleaf still using LTW files for Meditech
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

            'move files in reprocess to output directory
            dirs = Directory.GetFiles(strReprocessDirectory, "LTW.*") 'Note: Cloverleaf still using LTW files for Meditech
            For Each dir In dirs
                thefile = New FileInfo(dir)
                thefile.CopyTo(strOutputDirectory & thefile.Name)
                thefile.Delete()
            Next

        Catch ex As Exception
            functionError = True
            gblLogString = gblLogString & "Meditech Process Timeout Error" & vbCrLf
            gblLogString = gblLogString & ex.Message & vbCrLf

        Finally

        End Try


    End Sub

    Public Sub writeTolog(ByVal strMsg As String, ByVal eventType As Integer)

        Dim file As System.IO.StreamWriter
        Dim tempLogFileName As String = strLogDirectory & "ProcessTimeouts_Log.txt"
        file = My.Computer.FileSystem.OpenTextFileWriter(tempLogFileName, True)
        file.WriteLine(DateTime.Now & " : " & strMsg)
        file.Close()
    End Sub

End Module
