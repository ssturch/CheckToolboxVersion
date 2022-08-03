Imports System
Imports System.IO
Imports System.Collections

Module CheckToolboxVersion

    Public PathToSaveSWRlog As String
    Public CustomerToolboxFolder As String
    Sub Main()
        Dim SLDEDBChk As Boolean
        Dim WriteToSWRlog As Object = My.Computer.FileSystem
        Dim ProgrammDataFolder As String = My.Application.Info.DirectoryPath & "\DataFolder\"


        'Получение папки для сохранения результата анализа Toolbox клиента

        Dim ValidPath As Boolean = False

        While Not ValidPath
            Console.WriteLine("Введите папку, куда будет сохраняться SWRLog файл с результатом анализа Toolbox")
            PathToSaveSWRLog = Console.ReadLine()
            If My.Computer.FileSystem.DirectoryExists(PathToSaveSWRLog) = True Then
                Console.WriteLine("Данная папка существует")
                ValidPath = True
            Else
                Console.WriteLine("Данная папка не существует, проверьте правильность ввода и введите снова")

            End If
        End While

        PathToSaveSWRLog = PathToSaveSWRLog & "\CheckToolboxVersion" & ".swrlog"

        'Получение папки для анализа Toolbox клиента
        ValidPath = False

        Dim IsToolboxPath As Boolean = False

        While Not ValidPath
            Console.WriteLine("Введите папку с Toolbox для его анализа")
            CustomerToolboxFolder = Console.ReadLine()
            If My.Computer.FileSystem.DirectoryExists(CustomerToolboxFolder) = True Then
                Console.WriteLine("Данная папка существует")
                ValidPath = True
            Else
                Console.WriteLine("Данная папка не существует, проверьте правильность ввода и введите снова")

            End If
        End While

        'Анализ папки на наличие файлов sldedb или mdb

        While Not IsToolboxPath
            If My.Computer.FileSystem.FileExists(CustomerToolboxFolder & "\lang\English\SWBrowser.mdb") = True Or My.Computer.FileSystem.FileExists(CustomerToolboxFolder & "\lang\English\SWBrowser.sldedb") = True Then
                Console.WriteLine("Данная папка является папкой Toolbox" & vbCrLf & vbCrLf)
                IsToolboxPath = True
            Else
                Console.WriteLine("Данная папка не содержит компоненты Toolbox, проверьте правильность ввода и введите снова")
            End If
        End While

        'Создание файла лога для записи результатов

        Console.WriteLine("Создание файла CheckToolboxVersion.swrlog")
        Console.WriteLine("Анализ:")

        My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Результат анализа Toolbox клиента:" & vbCrLf, True)

        'Проверка, есть ли .sldedb файл

        If My.Computer.FileSystem.FileExists(CustomerToolboxFolder & "\lang\English\SWBrowser.sldedb") = True Then
            Console.WriteLine("Наличие .sldedb = OK")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Наличие .sldedb = OK" & vbCrLf, True)
            SLDEDBChk = True
        Else
            Console.WriteLine("Наличие .sldedb = NO")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Наличие .sldedb = NO" & vbCrLf, True)
        End If

        Call ToolboxVersionDATFilesCheck(CustomerToolboxFolder)
        Call CheckToolboxGOSTFoulderContent(ProgrammDataFolder)
        Call CheckToolboxGOSTFileContent(ProgrammDataFolder)
        Console.WriteLine("Проверка закончена, файл создан. Изучите полученные данные и нажмите любую кнопку")
        Console.ReadLine()

    End Sub

    Function CheckToolboxGOSTFoulderContent(ProgrammDataFolder As String)

        Dim StartCheckVersion As Boolean = False
        Dim ComparedPath As String
        Dim NewCheck As Boolean
        Dim ContinueCheck As Boolean = True
        Dim ChkIntgrt As Boolean = False
        Dim ActualFolders As New List(Of String)
        Dim CustomerFolders As New List(Of String)
        Dim FirstComparedToolboxVer As String
        Dim OlderComparedToolboxVer As String
        Dim FunctRepeat As Boolean = False
        Dim Stream As StreamReader = New StreamReader(ProgrammDataFolder & "\ActualToolboxFolderList.swrinstr")

        Console.WriteLine("==Проверка на соответствии версии Toolbox по папкам:==")
        My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "==Проверка на соответствии версии Toolbox по папкам:==" & vbCrLf, True)




        Do While Stream.Peek > -1
            Dim RLine As String = Stream.ReadLine()
            Select Case True
                Case RLine.Contains("START VERSION CHECK")
                    StartCheckVersion = True
                    Continue Do
                Case RLine.Contains("TOOLBOX") And StartCheckVersion And ContinueCheck
                    NewCheck = True
                    'FirstComparedToolboxVer = RLine.Replace(vbTab, "").Remove(FirstComparedToolboxVer.IndexOf("TO")).Replace("TOOLBOX ", "") ПОЧЕМУ ТО НЕ ОТРАБАТЫВАЕТ
                    'FirstComparedToolboxVer = RLine.Replace(vbTab, "") - данное выражение возвращает значение FIND, как будто полностью заменяет
                    'Поэтому пришлось городить дикий колхоз
                    FirstComparedToolboxVer = RLine
                    OlderComparedToolboxVer = RLine
                    FirstComparedToolboxVer = FirstComparedToolboxVer.Replace(vbTab, "").Replace("TOOLBOX ", "")
                    FirstComparedToolboxVer = FirstComparedToolboxVer.Remove(FirstComparedToolboxVer.IndexOf(" TO"))
                    OlderComparedToolboxVer = OlderComparedToolboxVer.Replace(vbTab, "").Replace("TOOLBOX " & FirstComparedToolboxVer & " TO ", "")
                    FunctRepeat = False
                    Continue Do
                Case RLine.Contains("END TOOLBOX")
                    NewCheck = False
                    Continue Do
                Case RLine.Contains("END VERSION CHECK")
                    StartCheckVersion = False
                    Continue Do
                Case RLine.Contains("START CHECK INTEGRITY")
                    StartCheckVersion = False
                    NewCheck = False
                    ContinueCheck = False
                    ChkIntgrt = True
                    Continue Do
                Case StartCheckVersion And NewCheck
                    Call PreliminaryFolderCheck(RLine, FirstComparedToolboxVer, OlderComparedToolboxVer, FunctRepeat)
                    FunctRepeat = True
                    Continue Do
                Case ChkIntgrt
                    ComparedPath = CustomerToolboxFolder & RLine
                    ActualFolders.Add(ComparedPath)
                    Continue Do
            End Select
        Loop

        For Each FoundFolder As String In My.Computer.FileSystem.GetDirectories(CustomerToolboxFolder & "\Browser\GOST", FileIO.SearchOption.SearchAllSubDirectories)
            CustomerFolders.Add(FoundFolder)
        Next

        For i = 0 To ActualFolders.Count - 1
            CustomerFolders.Remove(ActualFolders(i))
        Next

        Console.WriteLine("==Проверка соответствия структуры папок с последней актуальной поставкой:==")
        My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "==Проверка соответствия структуры папок с последней актуальной поставкой:==" & vbCrLf, True)

        If CustomerFolders.Count = 0 Then
            Console.WriteLine("Структура полностью совпадает с поставками Toolbox от SWR")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Структура полностью совпадает с поставками Toolbox от SWR" & vbCrLf, True)
        Else
            Console.WriteLine("Обнаружены папки, не входящие в поставку Toolbox от SWR:")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Обнаружены папки, не входящие в поставку Toolbox от SWR:" & vbCrLf, True)
            For i = 0 To CustomerFolders.Count - 1
                Console.WriteLine(CustomerFolders(i))
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, CustomerFolders(i) & vbCrLf, True)
            Next
        End If
    End Function
    Function PreliminaryFolderCheck(Wline, FirstComparedToolboxVer, OlderComparedToolboxVer, FunctRepeat)

        Dim ComparedPath As String


        ComparedPath = CustomerToolboxFolder & Wline.Replace(vbTab, "")


        If OlderComparedToolboxVer <> "2011" Then
            If FunctRepeat = False Then
                Console.WriteLine("Предварительная проверка на соответствие с версией от " & FirstComparedToolboxVer & " до " & OlderComparedToolboxVer)
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Предварительная проверка на соответствие с версией от " & FirstComparedToolboxVer & " до " & OlderComparedToolboxVer & vbCrLf, True)
            End If
            If My.Computer.FileSystem.DirectoryExists(ComparedPath) = True Then
                Console.WriteLine(Wline.Replace(vbTab, "") & " = ОК")
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, Wline & " = ОК" & vbCrLf, True)
            Else
                Console.WriteLine(Wline.Replace(vbTab, "") & " = NO")
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, Wline & " = NO" & vbCrLf, True)
            End If
        Else
            If FunctRepeat = False Then
                Console.WriteLine("Присутcтвие признаков версии Toolbox старше 2012.2")
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Присутcтвие признаков версии Toolbox старше 2012.2" & vbCrLf, True)
            End If
            If My.Computer.FileSystem.DirectoryExists(ComparedPath) = True Then
                Console.WriteLine(Wline.Replace(vbTab, "") & " = ОК")
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, Wline & " = ОК" & vbCrLf, True)
            Else
                Console.WriteLine(Wline.Replace(vbTab, "") & " = NO")
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, Wline & " = NO" & vbCrLf, True)
            End If
        End If
    End Function
    Function CheckToolboxGOSTFileContent(ProgrammDataFolder As String)

        Dim ToolboxVersion As String
        Dim StartCheck As Boolean = False
        Dim StartInstrNEW As Boolean = False
        Dim StartInstrDEL As Boolean = False
        Dim StartVerChk As Boolean = False

        Dim Stream As StreamReader = New StreamReader(ProgrammDataFolder & "\CompareList.swrinstr")


        Do While Stream.Peek > -1

            Dim WLine As String = Stream.ReadLine()


            Select Case True
                Case WLine.Contains("START VERSION CHECK")
                    StartCheck = True
                    ToolboxVersion = WLine.Remove(WLine.IndexOf("TO")).Replace("START VERSION CHECK ", "")
                    Console.WriteLine("  ==Проверка на соответствии версии Toolbox " & ToolboxVersion & ":==  ")
                    My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "  ==Проверка на соответствии версии Toolbox " & ToolboxVersion & ":==  " & vbCrLf, True)
                    Continue Do
                Case WLine.Contains("START NEW") And StartCheck
                    Console.WriteLine("Новые файлы, добавленные в этой версии:")
                    My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Новые файлы, добавленные в этой версии:" & vbCrLf, True)
                    StartInstrNEW = True
                    Continue Do
                Case WLine.Contains("END NEW")
                    StartInstrNEW = False
                    Continue Do
                Case WLine.Contains("START DELETED") And StartCheck
                    Console.WriteLine("Неактуальные мастер модели:")
                    My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Неактуальные мастер модели:" & vbCrLf, True)
                    StartInstrDEL = True
                    Continue Do
                Case WLine.Contains("END DELETED")
                    StartInstrDEL = False
                    Continue Do
                Case WLine.Contains("END VERSION CHECK")
                    StartCheck = False
                    Continue Do
                Case StartCheck And StartInstrNEW
                    Call CheckNewparts(WLine.Replace(vbTab & vbTab, ""))
                    Continue Do
                Case StartCheck And StartInstrDEL
                    Call CheckDeletedParts(WLine.Replace(vbTab & vbTab, ""))
                    Continue Do


            End Select
        Loop

    End Function
    Function CheckNewparts(WLine)

        Dim CharNumb As Integer
        Dim CmprFileCheckFolder As String
        Dim ComparedFilename As String
        Dim ComparedPath As String
        Dim OtherFilePath As New List(Of String)
        Dim OtherFilePathText As String

        ComparedPath = CustomerToolboxFolder & WLine
        CharNumb = InStrRev(ComparedPath, "\")
        ComparedFilename = ComparedPath.Remove(0, CharNumb)

        If My.Computer.FileSystem.FileExists(ComparedPath) = True Then
            Console.WriteLine(ComparedFilename & " = OK")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, ComparedFilename & " = OK" & vbCrLf, True)

            For Each FoundFile In My.Computer.FileSystem.GetFiles(CustomerToolboxFolder, FileIO.SearchOption.SearchAllSubDirectories, ComparedFilename)
                CmprFileCheckFolder = FoundFile
                If CmprFileCheckFolder <> ComparedPath Then
                    OtherFilePath.Add(CmprFileCheckFolder)
                End If
            Next

            If OtherFilePath.Count > 0 Then
                OtherFilePathText = String.Join(vbCrLf, OtherFilePath)
                Console.WriteLine("В базе содержатся копии " & ComparedFilename & " в папках:")
                Console.WriteLine(OtherFilePathText)

                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "В базе содержатся копии " & ComparedFilename & " в папках:" & vbCrLf, True)
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, OtherFilePathText & vbCrLf, True)

            End If

        Else

            For Each FoundFile In My.Computer.FileSystem.GetFiles(CustomerToolboxFolder, FileIO.SearchOption.SearchAllSubDirectories, ComparedFilename)
                CmprFileCheckFolder = FoundFile
            Next
            If CmprFileCheckFolder = Nothing Then
                Console.WriteLine(ComparedFilename & " = NO")
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, ComparedFilename & " = NO" & vbCrLf, True)
                CmprFileCheckFolder = Nothing
            Else
                If CmprFileCheckFolder <> ComparedPath Then
                    Console.WriteLine("Файл " & ComparedFilename & " был перемещен в папку " & CmprFileCheckFolder)
                    My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Файл " & ComparedFilename & " был перемещен в папку " & CmprFileCheckFolder & vbCrLf, True)
                End If

            End If


        End If


        CmprFileCheckFolder = Nothing
        ComparedFilename = Nothing

    End Function
    Function CheckDeletedParts(WLine)

        Dim CharNumb As Integer
        Dim CmprFileCheckFolder As String
        Dim ComparedFilename As String
        Dim ComparedPath As String
        Dim OtherFilePath As New List(Of String)
        Dim OtherFilePathText As String

        ComparedPath = CustomerToolboxFolder & WLine
        CharNumb = InStrRev(ComparedPath, "\")
        ComparedFilename = ComparedPath.Remove(0, CharNumb)

        If My.Computer.FileSystem.FileExists(ComparedPath) = True Then
            Console.WriteLine(ComparedFilename & " = OK")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, ComparedFilename & " = OK" & vbCrLf, True)

            For Each FoundFile In My.Computer.FileSystem.GetFiles(CustomerToolboxFolder, FileIO.SearchOption.SearchAllSubDirectories, ComparedFilename)
                CmprFileCheckFolder = FoundFile
                If CmprFileCheckFolder <> ComparedPath Then
                    OtherFilePath.Add(CmprFileCheckFolder)
                End If
            Next

            If OtherFilePath.Count > 0 Then
                OtherFilePathText = String.Join(vbCrLf, OtherFilePath)
                Console.WriteLine("В базе содержатся копии " & ComparedFilename & " в папках:")
                Console.WriteLine(OtherFilePathText)

                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "В базе содержатся копии " & ComparedFilename & " в папках:" & vbCrLf, True)
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, OtherFilePathText & vbCrLf, True)

            End If

        Else

            For Each FoundFile In My.Computer.FileSystem.GetFiles(CustomerToolboxFolder, FileIO.SearchOption.SearchAllSubDirectories, ComparedFilename)
                CmprFileCheckFolder = FoundFile
            Next
            If CmprFileCheckFolder = Nothing Then
                Console.WriteLine(ComparedFilename & " = NO")
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, ComparedFilename & " = NO" & vbCrLf, True)
                CmprFileCheckFolder = Nothing
            Else
                If CmprFileCheckFolder <> ComparedPath Then
                    Console.WriteLine("Файл " & ComparedFilename & " был перемещен в папку " & CmprFileCheckFolder)
                    My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Файл " & ComparedFilename & " был перемещен в папку " & CmprFileCheckFolder & vbCrLf, True)
                End If

            End If


        End If


        CmprFileCheckFolder = Nothing
        ComparedFilename = Nothing

    End Function
    Function ToolboxVersionDATFilesCheck(CustomerToolboxFolder)
        Dim ToolboxVersionTxtArr() As String
        Dim ToolboxVersion_GOSTTxtArr() As String

        Console.WriteLine("==Проверка на соответствии версии Toolbox по файлам ToolboxVersion.dat и ToolboxVersion_GOST.dat:==")
        My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "==Проверка на соответствии версии Toolbox по файлам ToolboxVersion.dat и ToolboxVersion_GOST.dat:==" & vbCrLf, True)

        If My.Computer.FileSystem.FileExists(CustomerToolboxFolder & "\ToolboxVersion.dat") = True Then
            Console.WriteLine("Наличие файла ToolboxVersion.dat = OK")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Наличие файла ToolboxVersion.dat = OK" & vbCrLf, True)
            ToolboxVersionTxtArr = Split(My.Computer.FileSystem.ReadAllText(CustomerToolboxFolder & "\ToolboxVersion.dat"), vbCrLf)
            Console.WriteLine("Версия оригинальной поставки - " & ToolboxVersionTxtArr(1))
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Версия оригинальной поставки - " & Replace(ToolboxVersionTxtArr(1), "// ", "") & vbCrLf, True)
        Else
            Console.WriteLine("Наличие файла ToolboxVersion.dat = NO")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Наличие файла ToolboxVersion.dat = NO" & vbCrLf, True)
        End If

        If My.Computer.FileSystem.FileExists(CustomerToolboxFolder & "\ToolboxVersion_GOST.dat") = True Then
            Console.WriteLine("Наличие файла ToolboxVersion_GOST.dat = OK")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Наличие файла ToolboxVersion_GOST.dat = OK" & vbCrLf, True)
            ToolboxVersion_GOSTTxtArr = Split(My.Computer.FileSystem.ReadAllText(CustomerToolboxFolder & "\ToolboxVersion_GOST.dat"), vbCrLf)
            Console.WriteLine("Версия оригинальной поставки - " & ToolboxVersion_GOSTTxtArr(1))
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Версия поставки от SWR - " & Replace(ToolboxVersion_GOSTTxtArr(1), "// ", "") & vbCrLf, True)
        Else
            Console.WriteLine("Наличие файла ToolboxVersion_GOST.dat = NO")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "Наличие файла ToolboxVersion_GOST.dat = NO" & vbCrLf, True)
        End If

    End Function



End Module
