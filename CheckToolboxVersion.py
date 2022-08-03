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


        '��������� ����� ��� ���������� ���������� ������� Toolbox �������

        Dim ValidPath As Boolean = False

        While Not ValidPath
            Console.WriteLine("������� �����, ���� ����� ����������� SWRLog ���� � ����������� ������� Toolbox")
            PathToSaveSWRLog = Console.ReadLine()
            If My.Computer.FileSystem.DirectoryExists(PathToSaveSWRLog) = True Then
                Console.WriteLine("������ ����� ����������")
                ValidPath = True
            Else
                Console.WriteLine("������ ����� �� ����������, ��������� ������������ ����� � ������� �����")

            End If
        End While

        PathToSaveSWRLog = PathToSaveSWRLog & "\CheckToolboxVersion" & ".swrlog"

        '��������� ����� ��� ������� Toolbox �������
        ValidPath = False

        Dim IsToolboxPath As Boolean = False

        While Not ValidPath
            Console.WriteLine("������� ����� � Toolbox ��� ��� �������")
            CustomerToolboxFolder = Console.ReadLine()
            If My.Computer.FileSystem.DirectoryExists(CustomerToolboxFolder) = True Then
                Console.WriteLine("������ ����� ����������")
                ValidPath = True
            Else
                Console.WriteLine("������ ����� �� ����������, ��������� ������������ ����� � ������� �����")

            End If
        End While

        '������ ����� �� ������� ������ sldedb ��� mdb

        While Not IsToolboxPath
            If My.Computer.FileSystem.FileExists(CustomerToolboxFolder & "\lang\English\SWBrowser.mdb") = True Or My.Computer.FileSystem.FileExists(CustomerToolboxFolder & "\lang\English\SWBrowser.sldedb") = True Then
                Console.WriteLine("������ ����� �������� ������ Toolbox" & vbCrLf & vbCrLf)
                IsToolboxPath = True
            Else
                Console.WriteLine("������ ����� �� �������� ���������� Toolbox, ��������� ������������ ����� � ������� �����")
            End If
        End While

        '�������� ����� ���� ��� ������ �����������

        Console.WriteLine("�������� ����� CheckToolboxVersion.swrlog")
        Console.WriteLine("������:")

        My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "��������� ������� Toolbox �������:" & vbCrLf, True)

        '��������, ���� �� .sldedb ����

        If My.Computer.FileSystem.FileExists(CustomerToolboxFolder & "\lang\English\SWBrowser.sldedb") = True Then
            Console.WriteLine("������� .sldedb = OK")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "������� .sldedb = OK" & vbCrLf, True)
            SLDEDBChk = True
        Else
            Console.WriteLine("������� .sldedb = NO")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "������� .sldedb = NO" & vbCrLf, True)
        End If

        Call ToolboxVersionDATFilesCheck(CustomerToolboxFolder)
        Call CheckToolboxGOSTFoulderContent(ProgrammDataFolder)
        Call CheckToolboxGOSTFileContent(ProgrammDataFolder)
        Console.WriteLine("�������� ���������, ���� ������. ������� ���������� ������ � ������� ����� ������")
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

        Console.WriteLine("==�������� �� ������������ ������ Toolbox �� ������:==")
        My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "==�������� �� ������������ ������ Toolbox �� ������:==" & vbCrLf, True)




        Do While Stream.Peek > -1
            Dim RLine As String = Stream.ReadLine()
            Select Case True
                Case RLine.Contains("START VERSION CHECK")
                    StartCheckVersion = True
                    Continue Do
                Case RLine.Contains("TOOLBOX") And StartCheckVersion And ContinueCheck
                    NewCheck = True
                    'FirstComparedToolboxVer = RLine.Replace(vbTab, "").Remove(FirstComparedToolboxVer.IndexOf("TO")).Replace("TOOLBOX ", "") ������ �� �� ������������
                    'FirstComparedToolboxVer = RLine.Replace(vbTab, "") - ������ ��������� ���������� �������� FIND, ��� ����� ��������� ��������
                    '������� �������� �������� ����� ������
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

        Console.WriteLine("==�������� ������������ ��������� ����� � ��������� ���������� ���������:==")
        My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "==�������� ������������ ��������� ����� � ��������� ���������� ���������:==" & vbCrLf, True)

        If CustomerFolders.Count = 0 Then
            Console.WriteLine("��������� ��������� ��������� � ���������� Toolbox �� SWR")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "��������� ��������� ��������� � ���������� Toolbox �� SWR" & vbCrLf, True)
        Else
            Console.WriteLine("���������� �����, �� �������� � �������� Toolbox �� SWR:")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "���������� �����, �� �������� � �������� Toolbox �� SWR:" & vbCrLf, True)
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
                Console.WriteLine("��������������� �������� �� ������������ � ������� �� " & FirstComparedToolboxVer & " �� " & OlderComparedToolboxVer)
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "��������������� �������� �� ������������ � ������� �� " & FirstComparedToolboxVer & " �� " & OlderComparedToolboxVer & vbCrLf, True)
            End If
            If My.Computer.FileSystem.DirectoryExists(ComparedPath) = True Then
                Console.WriteLine(Wline.Replace(vbTab, "") & " = ��")
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, Wline & " = ��" & vbCrLf, True)
            Else
                Console.WriteLine(Wline.Replace(vbTab, "") & " = NO")
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, Wline & " = NO" & vbCrLf, True)
            End If
        Else
            If FunctRepeat = False Then
                Console.WriteLine("������c���� ��������� ������ Toolbox ������ 2012.2")
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "������c���� ��������� ������ Toolbox ������ 2012.2" & vbCrLf, True)
            End If
            If My.Computer.FileSystem.DirectoryExists(ComparedPath) = True Then
                Console.WriteLine(Wline.Replace(vbTab, "") & " = ��")
                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, Wline & " = ��" & vbCrLf, True)
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
                    Console.WriteLine("  ==�������� �� ������������ ������ Toolbox " & ToolboxVersion & ":==  ")
                    My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "  ==�������� �� ������������ ������ Toolbox " & ToolboxVersion & ":==  " & vbCrLf, True)
                    Continue Do
                Case WLine.Contains("START NEW") And StartCheck
                    Console.WriteLine("����� �����, ����������� � ���� ������:")
                    My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "����� �����, ����������� � ���� ������:" & vbCrLf, True)
                    StartInstrNEW = True
                    Continue Do
                Case WLine.Contains("END NEW")
                    StartInstrNEW = False
                    Continue Do
                Case WLine.Contains("START DELETED") And StartCheck
                    Console.WriteLine("������������ ������ ������:")
                    My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "������������ ������ ������:" & vbCrLf, True)
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
                Console.WriteLine("� ���� ���������� ����� " & ComparedFilename & " � ������:")
                Console.WriteLine(OtherFilePathText)

                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "� ���� ���������� ����� " & ComparedFilename & " � ������:" & vbCrLf, True)
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
                    Console.WriteLine("���� " & ComparedFilename & " ��� ��������� � ����� " & CmprFileCheckFolder)
                    My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "���� " & ComparedFilename & " ��� ��������� � ����� " & CmprFileCheckFolder & vbCrLf, True)
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
                Console.WriteLine("� ���� ���������� ����� " & ComparedFilename & " � ������:")
                Console.WriteLine(OtherFilePathText)

                My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "� ���� ���������� ����� " & ComparedFilename & " � ������:" & vbCrLf, True)
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
                    Console.WriteLine("���� " & ComparedFilename & " ��� ��������� � ����� " & CmprFileCheckFolder)
                    My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "���� " & ComparedFilename & " ��� ��������� � ����� " & CmprFileCheckFolder & vbCrLf, True)
                End If

            End If


        End If


        CmprFileCheckFolder = Nothing
        ComparedFilename = Nothing

    End Function
    Function ToolboxVersionDATFilesCheck(CustomerToolboxFolder)
        Dim ToolboxVersionTxtArr() As String
        Dim ToolboxVersion_GOSTTxtArr() As String

        Console.WriteLine("==�������� �� ������������ ������ Toolbox �� ������ ToolboxVersion.dat � ToolboxVersion_GOST.dat:==")
        My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "==�������� �� ������������ ������ Toolbox �� ������ ToolboxVersion.dat � ToolboxVersion_GOST.dat:==" & vbCrLf, True)

        If My.Computer.FileSystem.FileExists(CustomerToolboxFolder & "\ToolboxVersion.dat") = True Then
            Console.WriteLine("������� ����� ToolboxVersion.dat = OK")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "������� ����� ToolboxVersion.dat = OK" & vbCrLf, True)
            ToolboxVersionTxtArr = Split(My.Computer.FileSystem.ReadAllText(CustomerToolboxFolder & "\ToolboxVersion.dat"), vbCrLf)
            Console.WriteLine("������ ������������ �������� - " & ToolboxVersionTxtArr(1))
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "������ ������������ �������� - " & Replace(ToolboxVersionTxtArr(1), "// ", "") & vbCrLf, True)
        Else
            Console.WriteLine("������� ����� ToolboxVersion.dat = NO")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "������� ����� ToolboxVersion.dat = NO" & vbCrLf, True)
        End If

        If My.Computer.FileSystem.FileExists(CustomerToolboxFolder & "\ToolboxVersion_GOST.dat") = True Then
            Console.WriteLine("������� ����� ToolboxVersion_GOST.dat = OK")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "������� ����� ToolboxVersion_GOST.dat = OK" & vbCrLf, True)
            ToolboxVersion_GOSTTxtArr = Split(My.Computer.FileSystem.ReadAllText(CustomerToolboxFolder & "\ToolboxVersion_GOST.dat"), vbCrLf)
            Console.WriteLine("������ ������������ �������� - " & ToolboxVersion_GOSTTxtArr(1))
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "������ �������� �� SWR - " & Replace(ToolboxVersion_GOSTTxtArr(1), "// ", "") & vbCrLf, True)
        Else
            Console.WriteLine("������� ����� ToolboxVersion_GOST.dat = NO")
            My.Computer.FileSystem.WriteAllText(PathToSaveSWRlog, "������� ����� ToolboxVersion_GOST.dat = NO" & vbCrLf, True)
        End If

    End Function



End Module
