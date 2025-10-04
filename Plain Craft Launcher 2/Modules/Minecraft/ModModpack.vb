Imports System.IO.Compression

Public Module ModModpack

    '触发整合包安装的外部接口
    ''' <summary>
    ''' 弹窗要求选择一个整合包文件并进行安装。
    ''' </summary>
    Public Sub ModpackInstall()
        Dim File As String = SelectFile("整合包文件(*.rar;*.zip;*.mrpack)|*.rar;*.zip;*.mrpack", "选择整合包压缩文件") '选择整合包文件
        If String.IsNullOrEmpty(File) Then Return
        RunInThread(
        Sub()
            Try
                ModpackInstall(File)
            Catch ex As CancelledException
            Catch ex As Exception
                Log(ex, "手动安装整合包失败", LogLevel.Msgbox)
            End Try
        End Sub)
    End Sub
    ''' <summary>
    ''' 构建并启动安装给定的整合包文件的加载器，并返回该加载器。若失败则抛出异常。
    ''' 必须在工作线程执行。
    ''' </summary>
    ''' <exception cref="CancelledException" />
    Public Function ModpackInstall(File As String, Optional VersionName As String = Nothing, Optional Logo As String = Nothing) As LoaderCombo(Of String)
        Log("[ModPack] 整合包安装请求：" & If(File, "null"))
        Dim Archive As ZipArchive = Nothing
        Dim ArchiveBaseFolder As String = ""
        Try
            '字符校验
            Dim TargetFolder As String = $"{PathMcFolder}versions\{VersionName}\"
            If TargetFolder.Contains("!") OrElse TargetFolder.Contains(";") Then Hint("游戏路径中不能含有感叹号或分号：" & TargetFolder, HintType.Critical) : Throw New CancelledException
            '获取整合包种类与关键 Json
            Dim PackType As Integer = -1
            Try
                Archive = New ZipArchive(New FileStream(File, FileMode.Open, FileAccess.Read, FileShare.Read))
                '从根目录判断整合包类型
                If Archive.GetEntry("mcbbs.packmeta") IsNot Nothing Then PackType = 3 : Exit Try 'MCBBS 整合包（优先于 manifest.json 判断）
                If Archive.GetEntry("mmc-pack.json") IsNot Nothing Then PackType = 2 : Exit Try 'MMC 整合包（优先于 manifest.json 判断，#4194）
                If Archive.GetEntry("modrinth.index.json") IsNot Nothing Then PackType = 4 : Exit Try 'Modrinth 整合包
                If Archive.GetEntry("manifest.json") IsNot Nothing Then
                    Dim Json As JObject = GetJson(ReadFile(Archive.GetEntry("manifest.json").Open, Encoding.UTF8))
                    If Json("addons") Is Nothing Then
                        PackType = 0 : Exit Try 'CurseForge 整合包
                    Else
                        PackType = 3 : Exit Try 'MCBBS 整合包
                    End If
                End If
                If Archive.GetEntry("modpack.json") IsNot Nothing Then PackType = 1 : Exit Try 'HMCL 整合包
                If Archive.GetEntry("modpack.zip") IsNot Nothing OrElse Archive.GetEntry("modpack.mrpack") IsNot Nothing Then PackType = 9 : Exit Try '带启动器的压缩包
                '从一级目录判断整合包类型
                For Each Entry In Archive.Entries
                    Dim FullNames As String() = Entry.FullName.Split("/")
                    ArchiveBaseFolder = FullNames(0) & "/"
                    '确定为一级目录下
                    If FullNames.Count <> 2 Then Continue For
                    '判断是否为关键文件
                    If FullNames(1) = "mcbbs.packmeta" Then PackType = 3 : Exit Try 'MCBBS 整合包（优先于 manifest.json 判断）
                    If FullNames(1) = "mmc-pack.json" Then PackType = 2 : Exit Try 'MMC 整合包（优先于 manifest.json 判断，#4194）
                    If FullNames(1) = "modrinth.index.json" Then PackType = 4 : Exit Try 'Modrinth 整合包
                    If FullNames(1) = "manifest.json" Then
                        Dim Json As JObject = GetJson(ReadFile(Entry.Open, Encoding.UTF8))
                        If Json("addons") Is Nothing Then
                            PackType = 0 : Exit Try 'CurseForge 整合包
                        Else
                            PackType = 3 : ArchiveBaseFolder = "overrides/" : Exit Try 'MCBBS 整合包
                        End If
                    End If
                    If FullNames(1) = "modpack.json" Then PackType = 1 : Exit Try 'HMCL 整合包
                    If FullNames(1) = "modpack.zip" OrElse FullNames(1) = "modpack.mrpack" Then PackType = 9 : Exit Try '带启动器的压缩包
                Next
            Catch ex As Exception
                If ex.GetDetail(True).Contains("Error.WinIOError") Then
                    Throw New Exception("打开整合包文件失败", ex)
                ElseIf File.EndsWithF(".rar", True) Then
                    Throw New Exception("PCL 无法处理 rar 格式的压缩包，请在解压后重新压缩为 zip 格式再试", ex)
                Else
                    Throw New Exception("打开整合包文件失败，文件可能损坏或为不支持的压缩包格式", ex)
                End If
            End Try
            '执行对应的安装方法
            Select Case PackType
                Case 0
                    Log("[ModPack] 整合包种类：CurseForge")
                    Return InstallPackCurseForge(File, Archive, ArchiveBaseFolder, VersionName, Logo)
                Case 1
                    Log("[ModPack] 整合包种类：HMCL")
                    Return InstallPackHMCL(File, Archive, ArchiveBaseFolder)
                Case 2
                    Log("[ModPack] 整合包种类：MMC")
                    Return InstallPackMMC(File, Archive, ArchiveBaseFolder)
                Case 3
                    Log("[ModPack] 整合包种类：MCBBS")
                    Return InstallPackMCBBS(File, Archive, ArchiveBaseFolder, VersionName)
                Case 4
                    Log("[ModPack] 整合包种类：Modrinth")
                    Return InstallPackModrinth(File, Archive, ArchiveBaseFolder, VersionName, Logo)
                Case 9
                    Log("[ModPack] 整合包种类：带启动器的压缩包")
                    Return InstallPackLauncherPack(File, Archive, ArchiveBaseFolder)
                Case Else
                    Log("[ModPack] 整合包种类：未能识别，假定为压缩包")
                    Return InstallPackCompress(File, Archive)
            End Select
        Finally
            If Archive IsNot Nothing Then Archive.Dispose()
        End Try
    End Function

    Private Sub ExtractModpackFiles(InstallTemp As String, FileAddress As String, Loader As LoaderBase, ProgressIncrement As Double)
        '解压文件
        Dim RetryCount As Integer = 1
        Dim Encode = Encoding.GetEncoding("GB18030")
        Dim InitialProgress = Loader.Progress
        Try
Retry:
            Loader.Progress = InitialProgress
            DeleteDirectory(InstallTemp)
            ExtractFile(FileAddress, InstallTemp, Encode, ProgressIncrementHandler:=Sub(Delta) Loader.Progress += Delta * ProgressIncrement)
        Catch ex As Exception
            Log(ex, "第 " & RetryCount & " 次解压尝试失败")
            If TypeOf ex Is ArgumentException Then
                Encode = Encoding.UTF8
                Log("[ModPack] 已切换压缩包解压编码为 UTF8")
            End If
            '完全不知道为啥会出现文件正在被另一进程使用的问题，总之加个重试
            If RetryCount < 5 Then
                Thread.Sleep(RetryCount * 2000)
                If Loader IsNot Nothing AndAlso Loader.LoadingState <> MyLoading.MyLoadingState.Run Then Return
                RetryCount += 1
                GoTo Retry
            Else
                Throw New Exception("解压整合包文件失败", ex)
            End If
        End Try
        Loader.Progress = InitialProgress + ProgressIncrement
    End Sub
    ''' <summary>
    ''' 从整合包的 override 目录复制文件，同时设置 PCL 的配置文件与版本隔离。
    ''' 对路径末尾是否为 \ 没有要求。
    ''' </summary>
    Private Sub CopyOverrideDirectory(OverridesFolder As String, VersionFolder As String, Loader As LoaderBase, ProgressIncrement As Double)
        If Not OverridesFolder.EndsWithF("\") Then OverridesFolder += "\"
        If Not VersionFolder.EndsWithF("\") Then VersionFolder += "\"
        '复制文件
        If Directory.Exists(OverridesFolder) Then
            Log($"[ModPack] 处理整合包覆写文件夹：{OverridesFolder} → {VersionFolder}")
            CopyDirectory(OverridesFolder, VersionFolder, Sub(Delta) Loader.Progress += Delta * ProgressIncrement)
        Else
            Log($"[ModPack] 整合包中没有覆写文件夹：{OverridesFolder}")
            Loader.Progress += ProgressIncrement
        End If
        '设置 ini
        Dim OverridesIni As String = $"{OverridesFolder}PCL\Setup.ini"
        Dim VersionIni As String = $"{VersionFolder}PCL\Setup.ini"
        If File.Exists(OverridesIni) Then
            WriteIni(OverridesIni, "VersionArgumentIndie", 1) '开启版本隔离
            WriteIni(OverridesIni, "VersionArgumentIndieV2", True)
            WriteIni(OverridesIni, "IsStar", False)
            CopyFile(OverridesIni, VersionIni) '覆写已有的 ini
        Else
            WriteIni(VersionIni, "VersionArgumentIndie", 1) '开启版本隔离
            WriteIni(VersionIni, "VersionArgumentIndieV2", True)
        End If
        IniClearCache(VersionIni) '重置缓存，避免被安装过程中写入的 ini 覆盖
    End Sub
    ''' <summary>
    ''' 弹窗提示整合包中存在不兼容的加载器。
    ''' 如果取消，则抛出 CancelledException。
    ''' </summary>
    ''' <exception cref="CancelledException" />
    Private Sub NotifyIncompatibleLoader(LoaderName As String)
        If MyMsgBox($"整合包中存在不兼容的加载器：{LoaderName}{vbCrLf}如果你知道如何手动安装它，可以先选择跳过，然后在整合包下载结束后手动安装。",
            "不兼容的加载器", "取消", $"不安装 {LoaderName} 并继续") = 1 Then Throw New CancelledException
    End Sub

#Region "不同类型整合包的安装方法"

    'CurseForge
    Private Function InstallPackCurseForge(FileAddress As String, Archive As Compression.ZipArchive, ArchiveBaseFolder As String,
                                           Optional VersionName As String = Nothing, Optional Logo As String = Nothing) As LoaderCombo(Of String)

        '读取 Json 文件
        Dim Json As JObject
        Try
            Json = GetJson(ReadFile(Archive.GetEntry(ArchiveBaseFolder & "manifest.json").Open))
        Catch ex As Exception
            Throw New Exception("CurseForge 整合包安装信息存在问题", ex)
        End Try
        If Json("minecraft") Is Nothing OrElse Json("minecraft")("version") Is Nothing Then Throw New Exception("CurseForge 整合包未提供 Minecraft 版本信息")

        '获取版本名
        If VersionName Is Nothing Then
            VersionName = If(Json("name"), "")
            Dim Validate As New ValidateFolderName(PathMcFolder & "versions")
            If Validate.Validate(VersionName) <> "" Then VersionName = ""
            If VersionName = "" Then VersionName = MyMsgBoxInput("输入版本名称", "", "", New ObjectModel.Collection(Of Validate) From {Validate})
            If String.IsNullOrEmpty(VersionName) Then Throw New CancelledException
        End If

        '检查并处理旧版本数据迁移
        VersionName = CheckAndMigrateOldVersion(VersionName)

        '获取 Mod API 版本信息
        Dim ForgeVersion As String = Nothing
        Dim NeoForgeVersion As String = Nothing
        Dim FabricVersion As String = Nothing
        For Each Entry In If(Json("minecraft")("modLoaders"), {})
            Dim Id As String = If(Entry("id"), "").ToString.ToLower
            If Id.StartsWithF("forge-") Then
                'Forge 指定
                If Id.Contains("recommended") Then Throw New Exception("该整合包版本过老，已不支持进行安装！")
                Log("[ModPack] 整合包 Forge 版本：" & Id)
                ForgeVersion = Id.Replace("forge-", "")
            ElseIf Id.StartsWithF("neoforge-") Then
                'NeoForge 指定
                Log("[ModPack] 整合包 NeoForge 版本：" & Id)
                NeoForgeVersion = Id.Replace("neoforge-", "")
            ElseIf Id.StartsWithF("fabric-") Then
                'Fabric 指定
                Log("[ModPack] 整合包 Fabric 版本：" & Id)
                FabricVersion = Id.Replace("fabric-", "")
            Else
                'ElseIf Id.StartsWithF("quilt-") Then
                NotifyIncompatibleLoader(Id)
            End If
        Next
        '解压
        Dim InstallTemp As String = RequestTaskTempFolder()
        Dim InstallLoaders As New List(Of LoaderBase)
        Dim OverrideHome As String = If(Json("overrides"), "")
        If OverrideHome <> "" Then
            InstallLoaders.Add(New LoaderTask(Of String, Integer)("解压整合包文件",
            Sub(Task As LoaderTask(Of String, Integer))
                ExtractModpackFiles(InstallTemp, FileAddress, Task, 0.6)
                CopyOverrideDirectory(
                    InstallTemp & ArchiveBaseFolder & If(OverrideHome = "." OrElse OverrideHome = "./", "", OverrideHome), '#5613
                    $"{PathMcFolder}versions\{VersionName}",
                    Task, 0.4)
            End Sub) With {
            .ProgressWeight = New FileInfo(FileAddress).Length / 1024 / 1024 / 6, .Block = False}) '每 6M 需要 1s
        End If
        '获取 Mod 列表
        Dim ModList As New List(Of Integer)
        Dim ModOptionalList As New List(Of Integer)
        For Each ModEntry In If(Json("files"), {})
            If ModEntry("projectID") Is Nothing OrElse ModEntry("fileID") Is Nothing Then
                Hint("某项 Mod 缺少必要信息，已跳过：" & ModEntry.ToString)
                Continue For
            End If
            ModList.Add(ModEntry("fileID"))
            If ModEntry("required") IsNot Nothing AndAlso Not ModEntry("required").ToObject(Of Boolean) Then ModOptionalList.Add(ModEntry("fileID"))
        Next
        If ModList.Any Then
            Dim ModDownloadLoaders As New List(Of LoaderBase)
            '获取 Mod 下载信息
            ModDownloadLoaders.Add(New LoaderTask(Of Integer, JArray)("获取 Mod 下载信息",
            Sub(Task As LoaderTask(Of Integer, JArray))
                Task.Output = DlModRequest("https://api.curseforge.com/v1/mods/files", HttpMethod.Post, "{""fileIds"": [" & Join(ModList, ",") & "]}", "application/json")("data")
                '如果文件已被删除，则 API 会跳过那一项
                If ModList.Count > Task.Output.Count Then Throw New Exception("整合包中的部分 Mod 版本已被 Mod 作者删除，所以没法继续安装了，请向整合包作者反馈该问题")
            End Sub) With {.ProgressWeight = ModList.Count / 10}) '每 10 Mod 需要 1s
            '构造 NetFile
            ModDownloadLoaders.Add(New LoaderTask(Of JArray, List(Of NetFile))("构造 Mod 下载信息",
            Sub(Task As LoaderTask(Of JArray, List(Of NetFile)))
                Dim FileList As New Dictionary(Of Integer, NetFile)
                For Each ModJson In Task.Input
                    Dim Id As Integer = ModJson("id").ToObject(Of Integer)
                    '跳过重复的 Mod（疑似 CurseForge Bug）
                    If FileList.ContainsKey(Id) Then Continue For
                    '可选 Mod 提示
                    If ModOptionalList.Contains(Id) Then
                        If MyMsgBox("是否要下载整合包中的可选文件 " & ModJson("displayName").ToString & "？", "下载可选文件", "是", "否") = 2 Then
                            Continue For
                        End If
                    End If
                    '根据 modules 和文件名后缀判断资源类型
                    Dim TargetFolder As String, Type As CompType
                    If ModJson("modules").Any Then 'modules 可能返回 null（#1006）
                        Dim ModuleNames = CType(ModJson("modules"), JArray).Select(Function(l) l("name").ToString).ToList
                        If ModuleNames.Contains("META-INF") OrElse ModuleNames.Contains("mcmod.info") OrElse
                           ModJson?("FileName")?.ToString?.EndsWithF(".jar", True) Then
                            TargetFolder = "mods" : Type = CompType.Mod
                        ElseIf ModuleNames.Contains("pack.mcmeta") Then
                            TargetFolder = "resourcepacks" : Type = CompType.ResourcePack
                        Else
                            TargetFolder = "shaderpacks" : Type = CompType.Shader
                        End If
                    Else
                        TargetFolder = "mods" : Type = CompType.Mod
                    End If
                    '建立 CompFile
                    Dim File As New CompFile(ModJson, Type)
                    If Not File.Available Then Continue For
                    '实际的添加
                    FileList.Add(Id, File.ToNetFile($"{PathMcFolder}versions\{VersionName}\{TargetFolder}\"))
                    Task.Progress += 1 / (1 + ModList.Count)
                Next
                Task.Output = FileList.Values.ToList
            End Sub) With {.ProgressWeight = ModList.Count / 200, .Show = False}) '每 200 Mod 需要 1s
            '下载 Mod 文件
            ModDownloadLoaders.Add(New LoaderDownload("下载 Mod", New List(Of NetFile)) With {.ProgressWeight = ModList.Count * 1.5}) '每个 Mod 需要 1.5s
            '构造加载器
            InstallLoaders.Add(New LoaderCombo(Of Integer)("下载 Mod（主加载器）", ModDownloadLoaders) With
                {.Show = False, .ProgressWeight = ModDownloadLoaders.Sum(Function(l) l.ProgressWeight)})
        End If

        '构造加载器
        Dim Request As New McInstallRequest With {
            .TargetVersionName = VersionName,
            .TargetVersionFolder = $"{PathMcFolder}versions\{VersionName}\",
            .MinecraftName = Json("minecraft")("version").ToString,
            .ForgeVersion = ForgeVersion,
            .NeoForgeVersion = NeoForgeVersion,
            .FabricVersion = FabricVersion
        }
        Dim MergeLoaders As List(Of LoaderBase) = McInstallLoader(Request)
        '构造总加载器
        Dim Loaders As New List(Of LoaderBase)
        Loaders.Add(New LoaderCombo(Of String)("整合包安装", InstallLoaders) With {.Show = False, .Block = False, .ProgressWeight = InstallLoaders.Sum(Function(l) l.ProgressWeight)})
        Loaders.Add(New LoaderCombo(Of String)("游戏安装", MergeLoaders) With {.Show = False, .ProgressWeight = MergeLoaders.Sum(Function(l) l.ProgressWeight)})
        Loaders.Add(New LoaderTask(Of String, String)("最终整理文件",
        Sub(Task As LoaderTask(Of String, String))
            '设置图标
            Dim VersionFolder As String = $"{PathMcFolder}versions\{VersionName}\"
            If Logo IsNot Nothing AndAlso File.Exists(Logo) Then
                File.Copy(Logo, VersionFolder & "PCL\Logo.png", True)
                WriteIni(VersionFolder & "PCL\Setup.ini", "Logo", "PCL\Logo.png")
                WriteIni(VersionFolder & "PCL\Setup.ini", "LogoCustom", "True")
                Log("[ModPack] 已设置整合包 Logo：" & Logo)
            End If
            '删除原始整合包文件
            For Each Target As String In {VersionFolder & "原始整合包.zip", VersionFolder & "原始整合包.mrpack"}
                If File.Exists(Target) Then
                    Log("[ModPack] 删除原始整合包文件：" & Target)
                    File.Delete(Target)
                End If
            Next
            If File.Exists(FileAddress) AndAlso IO.Path.GetFileNameWithoutExtension(FileAddress) = "modpack" Then
                Log("[ModPack] 删除安装整合包文件：" & FileAddress)
                File.Delete(FileAddress)
            End If
        End Sub) With {.ProgressWeight = 0.1, .Show = False})

        '重复任务检查
        Dim LoaderName As String = "CurseForge 整合包安装：" & VersionName & " "
        If LoaderTaskbar.Any(Function(l) l.Name = LoaderName) Then
            Hint("该整合包正在安装中！", HintType.Critical)
            Throw New CancelledException
        End If

        '启动
        Dim Loader As New LoaderCombo(Of String)(LoaderName, Loaders) With {.OnStateChanged = AddressOf McInstallState}
        Loader.Start(Request.TargetVersionFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        RunInUi(Sub() FrmMain.PageChange(FormMain.PageType.DownloadManager))
        Return Loader
    End Function

    'Modrinth
    Private Function InstallPackModrinth(FileAddress As String, Archive As Compression.ZipArchive, ArchiveBaseFolder As String, Optional VersionName As String = Nothing, Optional Logo As String = Nothing) As LoaderCombo(Of String)

        '读取 Json 文件
        Dim Json As JObject
        Try
            Json = GetJson(ReadFile(Archive.GetEntry(ArchiveBaseFolder & "modrinth.index.json").Open))
        Catch ex As Exception
            Throw New Exception("Modrinth 整合包安装信息存在问题", ex)
        End Try
        If Json("dependencies") Is Nothing OrElse Json("dependencies")("minecraft") Is Nothing Then Throw New Exception("Modrinth 整合包未提供 Minecraft 版本信息")
        '获取 Mod API 版本信息
        Dim MinecraftVersion As String = Nothing
        Dim ForgeVersion As String = Nothing
        Dim NeoForgeVersion As String = Nothing
        Dim FabricVersion As String = Nothing
        For Each Entry As JProperty In If(Json("dependencies"), {})
            Select Case Entry.Name.ToLower
                Case "minecraft"
                    MinecraftVersion = Entry.Value.ToString
                Case "forge" 'eg. 14.23.5.2859 / 1.19-41.1.0
                    ForgeVersion = Entry.Value.ToString
                    Log("[ModPack] 整合包 Forge 版本：" & ForgeVersion)
                Case "neoforge", "neo-forge" 'eg. 20.6.98-beta
                    NeoForgeVersion = Entry.Value.ToString
                    Log("[ModPack] 整合包 NeoForge 版本：" & NeoForgeVersion)
                Case "fabric-loader" 'eg. 0.14.14
                    FabricVersion = Entry.Value.ToString
                    Log("[ModPack] 整合包 Fabric 版本：" & FabricVersion)
                Case Else
                    'Case "quilt-loader" 'eg. 1.0.0
                    NotifyIncompatibleLoader(Entry.Name)
            End Select
        Next
        '获取版本名
        If VersionName Is Nothing Then
            VersionName = If(Json("name"), "")
            Dim Validate As New ValidateFolderName(PathMcFolder & "versions")
            If Validate.Validate(VersionName) <> "" Then VersionName = ""
            If VersionName = "" Then VersionName = MyMsgBoxInput("输入版本名称", "", "", New ObjectModel.Collection(Of Validate) From {Validate})
            If String.IsNullOrEmpty(VersionName) Then Throw New CancelledException
        End If

        '检查并处理旧版本数据迁移
        VersionName = CheckAndMigrateOldVersion(VersionName)

        '解压
        Dim InstallTemp As String = RequestTaskTempFolder()
        Dim InstallLoaders As New List(Of LoaderBase)
        InstallLoaders.Add(New LoaderTask(Of String, Integer)("解压整合包文件",
        Sub(Task As LoaderTask(Of String, Integer))
            ExtractModpackFiles(InstallTemp, FileAddress, Task, 0.5)
            CopyOverrideDirectory(
                InstallTemp & ArchiveBaseFolder & "overrides",
                PathMcFolder & "versions\" & VersionName,
                Task, 0.4)
            CopyOverrideDirectory(
                InstallTemp & ArchiveBaseFolder & "client-overrides",
                PathMcFolder & "versions\" & VersionName,
                Task, 0.1)
        End Sub) With {.ProgressWeight = New FileInfo(FileAddress).Length / 1024 / 1024 / 6, .Block = False}) '每 6M 需要 1s
        '获取下载文件列表
        Dim FileList As New List(Of NetFile)
        For Each File In If(Json("files"), {})
            '检查是否需要该文件
            If File("env") IsNot Nothing Then
                Select Case File("env")("client").ToString
                    Case "optional"
                        If MyMsgBox("是否要下载可选文件 " & GetFileNameFromPath(File("path").ToString) & "？",
                                    "下载可选文件", "是", "否") = 2 Then
                            Continue For
                        End If
                    Case "unsupported"
                        Continue For
                End Select
            End If
            '添加下载文件
            Dim Urls = File("downloads").SelectMany(Function(t) CompFile.HandleCurseForgeDownloadUrls(t.ToString)).ToList
            Urls.AddRange(Urls.Select(Function(u) DlSourceModGet(u)).ToList)
            Urls = Urls.Distinct.ToList()
            Dim TargetPath As String = $"{PathMcFolder}versions\{VersionName}\{File("path")}"
            If Not IO.Path.GetFullPath(TargetPath).StartsWithF($"{PathMcFolder}versions\{VersionName}\", True) Then
                MyMsgBox($"整合包的文件路径超出了版本文件夹，请向整合包作者反馈此问题！" & vbCrLf & "错误的文件：" & TargetPath, "文件路径校验失败", IsWarn:=True)
                Throw New CancelledException
            End If
            FileList.Add(New NetFile(Urls, TargetPath,
                New FileChecker(ActualSize:=File("fileSize").ToObject(Of Long), Hash:=File("hashes")("sha1").ToString), True))
        Next
        If FileList.Any Then
            InstallLoaders.Add(New LoaderDownload("下载额外文件", FileList) With {.ProgressWeight = FileList.Count * 1.5}) '每个 Mod 需要 1.5s
        End If

        '构造加载器
        Dim Request As New McInstallRequest With {
            .TargetVersionName = VersionName,
            .TargetVersionFolder = $"{PathMcFolder}versions\{VersionName}\",
            .MinecraftName = MinecraftVersion,
            .ForgeVersion = ForgeVersion,
            .NeoForgeVersion = NeoForgeVersion,
            .FabricVersion = FabricVersion
        }
        Dim MergeLoaders As List(Of LoaderBase) = McInstallLoader(Request)
        '构造总加载器
        Dim Loaders As New List(Of LoaderBase)
        Loaders.Add(New LoaderCombo(Of String)("整合包安装", InstallLoaders) With {.Show = False, .Block = False, .ProgressWeight = InstallLoaders.Sum(Function(l) l.ProgressWeight)})
        Loaders.Add(New LoaderCombo(Of String)("游戏安装", MergeLoaders) With {.Show = False, .ProgressWeight = MergeLoaders.Sum(Function(l) l.ProgressWeight)})
        Loaders.Add(New LoaderTask(Of String, String)("最终整理文件",
        Sub(Task As LoaderTask(Of String, String))
            '设置图标
            Dim VersionFolder As String = $"{PathMcFolder}versions\{VersionName}\"
            If Logo IsNot Nothing AndAlso File.Exists(Logo) Then
                File.Copy(Logo, VersionFolder & "PCL\Logo.png", True)
                WriteIni(VersionFolder & "PCL\Setup.ini", "Logo", "PCL\Logo.png")
                WriteIni(VersionFolder & "PCL\Setup.ini", "LogoCustom", "True")
                Log("[ModPack] 已设置整合包 Logo：" & Logo)
            End If
            '删除原始整合包文件
            For Each Target As String In {VersionFolder & "原始整合包.zip", VersionFolder & "原始整合包.mrpack"}
                If File.Exists(Target) Then
                    Log("[ModPack] 删除原始整合包文件：" & Target)
                    File.Delete(Target)
                End If
            Next
            If File.Exists(FileAddress) AndAlso IO.Path.GetFileNameWithoutExtension(FileAddress) = "modpack" Then
                Log("[ModPack] 删除安装整合包文件：" & FileAddress)
                File.Delete(FileAddress)
            End If
        End Sub) With {.ProgressWeight = 0.1, .Show = False})

        '重复任务检查
        Dim LoaderName As String = $"Modrinth 整合包安装：{VersionName} "
        If LoaderTaskbar.Any(Function(l) l.Name = LoaderName) Then
            Hint("该整合包正在安装中！", HintType.Critical)
            Throw New CancelledException
        End If

        '启动
        Dim Loader As New LoaderCombo(Of String)(LoaderName, Loaders) With {.OnStateChanged = AddressOf McInstallState}
        Loader.Start(Request.TargetVersionFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        RunInUi(Sub() FrmMain.PageChange(FormMain.PageType.DownloadManager))
        Return Loader
    End Function

    'HMCL
    Private Function InstallPackHMCL(FileAddress As String, Archive As Compression.ZipArchive, ArchiveBaseFolder As String) As LoaderCombo(Of String)
        '读取 Json 文件
        Dim Json As JObject
        Try
            Json = GetJson(ReadFile(Archive.GetEntry(ArchiveBaseFolder & "modpack.json").Open, Encoding.UTF8))
        Catch ex As Exception
            Throw New Exception("HMCL 整合包安装信息存在问题", ex)
        End Try
        '获取版本名
        Dim VersionName As String = If(Json("name"), "")
        Dim Validate As New ValidateFolderName(PathMcFolder & "versions")
        If Validate.Validate(VersionName) <> "" Then VersionName = ""
        If VersionName = "" Then VersionName = MyMsgBoxInput("输入版本名称", "", "", New ObjectModel.Collection(Of Validate) From {Validate})
        If String.IsNullOrEmpty(VersionName) Then Throw New CancelledException
        '解压
        Dim InstallTemp As String = RequestTaskTempFolder()
        Dim InstallLoaders As New List(Of LoaderBase)
        InstallLoaders.Add(New LoaderTask(Of String, Integer)("解压整合包文件",
        Sub(Task As LoaderTask(Of String, Integer))
            ExtractModpackFiles(InstallTemp, FileAddress, Task, 0.6)
            CopyOverrideDirectory(
                InstallTemp & ArchiveBaseFolder & "minecraft",
                PathMcFolder & "versions\" & VersionName,
                Task, 0.4)
        End Sub) With {.ProgressWeight = New FileInfo(FileAddress).Length / 1024 / 1024 / 6, .Block = False}) '每 6M 需要 1s
        '构造游戏本体安装加载器
        If Json("gameVersion") Is Nothing Then Throw New Exception("该 HMCL 整合包未提供游戏版本信息，无法安装！")
        Dim Request As New McInstallRequest With {
            .TargetVersionName = VersionName,
            .TargetVersionFolder = $"{PathMcFolder}versions\{VersionName}\",
            .MinecraftName = Json("gameVersion").ToString
        }
        Dim MergeLoaders As List(Of LoaderBase) = McInstallLoader(Request)
        '构造总加载器
        Dim Loaders As New List(Of LoaderBase) From {
            New LoaderCombo(Of String)("整合包安装", InstallLoaders) With {.Show = False, .Block = False, .ProgressWeight = InstallLoaders.Sum(Function(l) l.ProgressWeight)},
            New LoaderCombo(Of String)("游戏安装", MergeLoaders) With {.Show = False, .ProgressWeight = MergeLoaders.Sum(Function(l) l.ProgressWeight)}
        }
        '重复任务检查
        Dim LoaderName As String = "HMCL 整合包安装：" & VersionName & " "
        If LoaderTaskbar.Any(Function(l) l.Name = LoaderName) Then
            Hint("该整合包正在安装中！", HintType.Critical)
            Throw New CancelledException
        End If
        '启动
        Dim Loader As New LoaderCombo(Of String)(LoaderName, Loaders) With {.OnStateChanged = AddressOf McInstallState}
        Loader.Start(Request.TargetVersionFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        RunInUi(Sub() FrmMain.PageChange(FormMain.PageType.DownloadManager))
        Return Loader
    End Function

    'MMC
    Private Function InstallPackMMC(FileAddress As String, Archive As Compression.ZipArchive, ArchiveBaseFolder As String) As LoaderCombo(Of String)
        '读取 Json 文件
        Dim PackJson As JObject, PackInstance As String
        Try
            PackJson = GetJson(ReadFile(Archive.GetEntry(ArchiveBaseFolder & "mmc-pack.json").Open, Encoding.UTF8))
            PackInstance = ReadFile(Archive.GetEntry(ArchiveBaseFolder & "instance.cfg").Open, Encoding.UTF8)
        Catch ex As Exception
            Throw New Exception("MMC 整合包安装信息存在问题", ex)
        End Try
        '获取版本名
        Dim VersionName As String = If(RegexSeek(PackInstance, "(?<=\nname\=)[^\n]+"), "")
        Dim Validate As New ValidateFolderName(PathMcFolder & "versions")
        If Validate.Validate(VersionName) <> "" Then VersionName = ""
        If VersionName = "" Then VersionName = MyMsgBoxInput("输入版本名称", "", "", New ObjectModel.Collection(Of Validate) From {Validate})
        If String.IsNullOrEmpty(VersionName) Then Throw New CancelledException
        '解压
        Dim InstallTemp As String = RequestTaskTempFolder()
        Dim SetupFile As String = $"{PathMcFolder}versions\{VersionName}\PCL\Setup.ini"
        Dim InstallLoaders As New List(Of LoaderBase)
        InstallLoaders.Add(New LoaderTask(Of String, Integer)("解压整合包文件",
        Sub(Task As LoaderTask(Of String, Integer))
            ExtractModpackFiles(InstallTemp, FileAddress, Task, 0.55)
            CopyOverrideDirectory(
                InstallTemp & ArchiveBaseFolder & ".minecraft",
                PathMcFolder & "versions\" & VersionName,
                Task, 0.4)
            '读取 MMC 设置文件（#2655）
            Try
                Dim MMCSetupFile As String = InstallTemp & ArchiveBaseFolder & "instance.cfg"
                If File.Exists(MMCSetupFile) Then
                    '将其中的等号替换为冒号，以符合 ini 文件格式
                    Dim Lines As New List(Of String)
                    For Each Line In ReadFile(MMCSetupFile).Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                        If Not Line.Contains("=") Then Continue For
                        Lines.Add(Line.BeforeFirst("=") & ":" & Line.AfterFirst("="))
                    Next
                    WriteFile(MMCSetupFile, Join(Lines, vbCrLf))
                    '读取文件
                    If ReadIni(MMCSetupFile, "OverrideCommands", False) Then
                        Dim PreLaunchCommand As String = ReadIni(MMCSetupFile, "PreLaunchCommand")
                        If PreLaunchCommand <> "" Then
                            PreLaunchCommand = PreLaunchCommand.Replace("\""", """").
                                Replace("$INST_JAVA", "{java}javaw.exe").
                                Replace("$INST_MC_DIR\", "{minecraft}").Replace("$INST_MC_DIR", "{minecraft}").
                                Replace("$INST_DIR\", "{verpath}").Replace("$INST_DIR", "{verpath}").
                                Replace("$INST_ID", "{name}").Replace("$INST_NAME", "{name}")
                            WriteIni(SetupFile, "VersionAdvanceRun", PreLaunchCommand)
                            Log("[ModPack] 迁移 MultiMC 版本独立设置：启动前执行命令：" & PreLaunchCommand)
                        End If
                    End If
                    If ReadIni(MMCSetupFile, "JoinServerOnLaunch", False) Then
                        Dim ServerAddress As String = ReadIni(MMCSetupFile, "JoinServerOnLaunchAddress").Replace("\""", """")
                        WriteIni(SetupFile, "VersionServerEnter", ServerAddress)
                        Log("[ModPack] 迁移 MultiMC 版本独立设置：自动进入服务器：" & ServerAddress)
                    End If
                    If ReadIni(MMCSetupFile, "IgnoreJavaCompatibility", False) Then
                        WriteIni(SetupFile, "VersionAdvanceJava", True)
                        Log("[ModPack] 迁移 MultiMC 版本独立设置：忽略 Java 兼容性警告")
                    End If
                    Dim Logo As String = ReadIni(MMCSetupFile, "iconKey", "")
                    If Logo <> "" AndAlso File.Exists($"{InstallTemp}{ArchiveBaseFolder}{Logo}.png") Then
                        WriteIni(SetupFile, "LogoCustom", True)
                        WriteIni(SetupFile, "Logo", "PCL\Logo.png")
                        CopyFile($"{InstallTemp}{ArchiveBaseFolder}{Logo}.png", $"{PathMcFolder}versions\{VersionName}\PCL\Logo.png")
                        Log($"[ModPack] 迁移 MultiMC 版本独立设置：版本图标（{Logo}.png）")
                    End If
                    'JVM 参数
                    Dim JvmArgs As String = ReadIni(MMCSetupFile, "JvmArgs", "")
                    If JvmArgs <> "" Then
                        If ReadIni(MMCSetupFile, "OverrideJavaArgs", False) Then
                            WriteIni(SetupFile, "VersionAdvanceJvm", JvmArgs)
                            Log("[ModPack] 迁移 MultiMC 版本独立设置：JVM 参数（覆盖）：" & JvmArgs)
                        Else
                            JvmArgs += " " & Setup.Get("LaunchAdvanceJvm")
                            WriteIni(SetupFile, "VersionAdvanceJvm", JvmArgs)
                            Log("[ModPack] 迁移 MultiMC 版本独立设置：JVM 参数（追加）：" & JvmArgs)
                        End If
                    End If
                End If
            Catch ex As Exception
                Log(ex, $"读取 MMC 配置文件失败（{InstallTemp}{ArchiveBaseFolder}instance.cfg）")
            End Try
        End Sub) With {.ProgressWeight = New FileInfo(FileAddress).Length / 1024 / 1024 / 6, .Block = False}) '每 6M 需要 1s
        '构造版本安装请求
        If PackJson("components") Is Nothing Then Throw New Exception("该 MMC 整合包未提供游戏版本信息，无法安装！")
        Dim Request As New McInstallRequest With {.TargetVersionName = VersionName, .TargetVersionFolder = $"{PathMcFolder}versions\{VersionName}\"}
        For Each Component In PackJson("components")
            If Not Component.Contains("uid") Then Continue For
            Select Case Component("uid").ToString
                Case "org.lwjgl"
                    Log("[ModPack] 已跳过 LWJGL 项")
                Case "net.minecraft"
                    Request.MinecraftName = Component("version")
                Case "net.minecraftforge"
                    Request.ForgeVersion = Component("version")
                Case "net.neoforged"
                    Request.NeoForgeVersion = Component("version")
                Case "net.fabricmc.fabric-loader"
                    Request.FabricVersion = Component("version")
                Case Else
                    'Case "org.quiltmc.quilt-loader" 'eg. 1.0.0
                    NotifyIncompatibleLoader(Component("uid"))
            End Select
        Next
        '构造加载器
        Dim MergeLoaders As List(Of LoaderBase) = McInstallLoader(Request)
        '构造总加载器
        Dim Loaders As New List(Of LoaderBase)
        Loaders.Add(New LoaderCombo(Of String)("整合包安装", InstallLoaders) With {.Show = False, .Block = False, .ProgressWeight = InstallLoaders.Sum(Function(l) l.ProgressWeight)})
        Loaders.Add(New LoaderCombo(Of String)("游戏安装", MergeLoaders) With {.Show = False, .ProgressWeight = MergeLoaders.Sum(Function(l) l.ProgressWeight)})

        '重复任务检查
        Dim LoaderName As String = "MMC 整合包安装：" & VersionName & " "
        If LoaderTaskbar.Any(Function(l) l.Name = LoaderName) Then
            Hint("该整合包正在安装中！", HintType.Critical)
            Throw New CancelledException
        End If

        '启动
        Dim Loader As New LoaderCombo(Of String)(LoaderName, Loaders) With {.OnStateChanged = AddressOf McInstallState}
        Loader.Start(Request.TargetVersionFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        RunInUi(Sub() FrmMain.PageChange(FormMain.PageType.DownloadManager))
        Return Loader
    End Function

    'MCBBS
    Private Function InstallPackMCBBS(FileAddress As String, Archive As Compression.ZipArchive, ArchiveBaseFolder As String,
                                      Optional VersionName As String = Nothing) As LoaderCombo(Of String)
        '读取 Json 文件
        Dim Json As JObject
        Try
            Dim Entry = If(Archive.GetEntry(ArchiveBaseFolder & "mcbbs.packmeta"), Archive.GetEntry(ArchiveBaseFolder & "manifest.json"))
            Json = GetJson(ReadFile(Entry.Open, Encoding.UTF8))
        Catch ex As Exception
            Throw New Exception("MCBBS 整合包安装信息存在问题", ex)
        End Try
        '获取版本名
        If VersionName Is Nothing Then
            VersionName = If(Json("name"), "")
            Dim Validate As New ValidateFolderName(PathMcFolder & "versions")
            If Validate.Validate(VersionName) <> "" Then VersionName = ""
            If VersionName = "" Then VersionName = MyMsgBoxInput("输入版本名称", "", "", New ObjectModel.Collection(Of Validate) From {Validate})
            If String.IsNullOrEmpty(VersionName) Then Throw New CancelledException
        End If

        '检查并处理旧版本数据迁移
        VersionName = CheckAndMigrateOldVersion(VersionName)

        '解压
        Dim InstallTemp As String = RequestTaskTempFolder()
        Dim SetupFile As String = $"{PathMcFolder}versions\{VersionName}\PCL\Setup.ini"
        Dim InstallLoaders As New List(Of LoaderBase)
        InstallLoaders.Add(New LoaderTask(Of String, Integer)("解压整合包文件",
        Sub(Task As LoaderTask(Of String, Integer))
            ExtractModpackFiles(InstallTemp, FileAddress, Task, 0.6)
            CopyOverrideDirectory(
                InstallTemp & ArchiveBaseFolder & "overrides",
                PathMcFolder & "versions\" & VersionName,
                Task, 0.4)
            'JVM 参数
            If Json("launchInfo") IsNot Nothing Then
                Dim LaunchInfo As JObject = Json("launchInfo")
                If LaunchInfo.ContainsKey("javaArgument") Then WriteIni(SetupFile, "VersionAdvanceJvm", String.Join(" ", LaunchInfo("javaArgument")))
                If LaunchInfo.ContainsKey("launchArgument") Then WriteIni(SetupFile, "VersionAdvanceGame", String.Join(" ", LaunchInfo("launchArgument")))
            End If
        End Sub) With {.ProgressWeight = New FileInfo(FileAddress).Length / 1024 / 1024 / 6, .Block = False}) '每 6M 需要 1s
        '构造加载器
        If Json("addons") Is Nothing Then Throw New Exception("该 MCBBS 整合包未提供游戏版本附加信息，无法安装！")
        Dim Addons As New Dictionary(Of String, String)
        For Each Entry In Json("addons")
            Addons.Add(Entry("id"), Entry("version"))
        Next
        If Not Addons.ContainsKey("game") Then Throw New Exception("该 MCBBS 整合包未提供游戏版本信息，无法安装！")
        If Addons.ContainsKey("quilt") Then
            Hint("PCL 暂不支持安装需要 Quilt 的整合包！", HintType.Critical)
            Throw New CancelledException
        End If
        Dim Request As New McInstallRequest With {
            .TargetVersionName = VersionName,
            .TargetVersionFolder = $"{PathMcFolder}versions\{VersionName}\",
            .MinecraftName = Addons("game"),
            .OptiFineVersion = If(Addons.ContainsKey("optifine"), Addons("optifine"), Nothing),
            .ForgeVersion = If(Addons.ContainsKey("forge"), Addons("forge"), Nothing),
            .NeoForgeVersion = If(Addons.ContainsKey("neoforge"), Addons("neoforge"), Nothing),
            .FabricVersion = If(Addons.ContainsKey("fabric"), Addons("fabric"), Nothing)
        }
        Dim MergeLoaders As List(Of LoaderBase) = McInstallLoader(Request)
        '构造总加载器
        Dim Loaders As New List(Of LoaderBase)
        Loaders.Add(New LoaderCombo(Of String)("整合包安装", InstallLoaders) With {.Show = False, .Block = False, .ProgressWeight = InstallLoaders.Sum(Function(l) l.ProgressWeight)})
        Loaders.Add(New LoaderCombo(Of String)("游戏安装", MergeLoaders) With {.Show = False, .ProgressWeight = MergeLoaders.Sum(Function(l) l.ProgressWeight)})

        '重复任务检查
        Dim LoaderName As String = "MCBBS 整合包安装：" & VersionName & " "
        If LoaderTaskbar.Any(Function(l) l.Name = LoaderName) Then
            Hint("该整合包正在安装中！", HintType.Critical)
            Throw New CancelledException
        End If

        '启动
        Dim Loader As New LoaderCombo(Of String)(LoaderName, Loaders) With {.OnStateChanged = AddressOf McInstallState}
        'If Archive IsNot Nothing Then Archive.Dispose() '解除占用，以免在加载器中触发 “正由另一进程使用，因此该进程无法访问此文件”
        Loader.Start(Request.TargetVersionFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        RunInUi(Sub() FrmMain.PageChange(FormMain.PageType.DownloadManager))
        Return Loader
    End Function

    '带启动器的压缩包
    Private Function InstallPackLauncherPack(FileAddress As String, Archive As Compression.ZipArchive, ArchiveBaseFolder As String) As LoaderCombo(Of String)
        '获取解压路径
        MyMsgBox("接下来请选择一个空文件夹，它会被安装到这个文件夹里。", "安装", "继续", ForceWait:=True)
        Dim TargetFolder As String = SelectFolder("选择安装目标（必须是一个空文件夹）")
        If String.IsNullOrEmpty(TargetFolder) Then Throw New CancelledException
        If Directory.GetFileSystemEntries(TargetFolder).Length > 0 Then Hint("请选择一个空文件夹作为安装目标！", HintType.Critical) : Throw New CancelledException
        '解压
        Dim Loader As New LoaderCombo(Of String)("解压压缩包", {
            New LoaderTask(Of String, Integer)("解压压缩包",
            Sub(Task As LoaderTask(Of String, Integer))
                ExtractModpackFiles(TargetFolder, FileAddress, Task, 0.9)
                Thread.Sleep(400) '避免文件争用
                '查找解压后的 exe 文件
                Dim Launcher As String = Nothing
                For Each ExeFile In Directory.GetFiles(TargetFolder, "*.exe", SearchOption.TopDirectoryOnly)
                    Dim Info = FileVersionInfo.GetVersionInfo(ExeFile)
                    Log($"[Modpack] 文件 {ExeFile} 的产品名标识为 {Info.ProductName}")
                    If Info.ProductName = "Plain Craft Launcher" Then
                        Launcher = ExeFile
                        Log($"[Modpack] 发现整合包附带的 PCL 启动器：{ExeFile}")
                    ElseIf (Info.ProductName.ContainsF("Launcher", True) OrElse Info.ProductName.ContainsF("启动", True)) AndAlso
                        Not Info.ProductName = "Plain Craft Launcher Admin Manager" Then
                        If Launcher Is Nothing Then
                            Launcher = ExeFile
                            Log($"[Modpack] 发现整合包附带的疑似第三方启动器：{ExeFile}")
                        End If
                    End If
                Next
                Task.Progress = 0.95
                '尝试使用附带的启动器打开
                If Launcher IsNot Nothing Then
                    Log("[Modpack] 找到压缩包中附带的启动器：" & Launcher)
                    If MyMsgBox($"整合包里似乎自带了启动器，是否换用它继续安装？{vbCrLf}即将打开：{Launcher}", "换用整合包启动器？", "换用", "不换用") = 1 Then
                        OpenExplorer(TargetFolder)
                        StartProcess(Launcher, "--wait") '要求等待已有的 PCL 退出
                        Log("[Modpack] 为换用整合包中的启动器启动，强制结束程序")
                        FrmMain.EndProgram(False)
                        Return
                    End If
                Else
                    Log("[Modpack] 未找到压缩包中附带的启动器")
                End If
                OpenExplorer(TargetFolder)
                '加入文件夹列表
                Dim VersionName As String = GetFolderNameFromPath(TargetFolder)
                Directory.CreateDirectory(TargetFolder & ".minecraft\")
                PageSelectLeft.AddFolder(
                    TargetFolder & ".minecraft\" & ArchiveBaseFolder.Replace("/", "\").TrimStart("\"), '格式例如：包裹文件夹\.minecraft\（最短为空字符串）
                    VersionName, False)
                '调用 modpack 文件进行安装
                Dim ModpackFile = Directory.GetFiles(TargetFolder, "modpack.*", SearchOption.AllDirectories).First
                Log("[Modpack] 调用 modpack 文件继续安装：" & ModpackFile)
                ModpackInstall(ModpackFile)
            End Sub)
        })
        Loader.Start(TargetFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        FrmMain.BtnExtraDownload.Ribble()
        Return Loader
    End Function

    '普通压缩包
    Private Function InstallPackCompress(FileAddress As String, Archive As Compression.ZipArchive) As LoaderCombo(Of String)
        '尝试定位 .minecraft 文件夹：寻找形如 “/versions/XXX/XXX.json” 的路径
        Dim Match As RegularExpressions.Match = Nothing
        Dim Regex As New RegularExpressions.Regex("^.*\/(?=versions\/(?<ver>[^\/]+)\/(\k<ver>)\.json$)", RegularExpressions.RegexOptions.IgnoreCase)
        For Each Entry In Archive.Entries
            Dim EntryMatch = Regex.Match("/" & Entry.FullName)
            If EntryMatch.Success Then
                Match = EntryMatch
                Exit For
            End If
        Next
        If Match Is Nothing Then Throw New Exception("未能找到适合的文件结构，这可能不是一个 MC 压缩包") '没有匹配
        Dim ArchiveBaseFolder As String = Match.Value.Replace("/", "\").TrimStart("\") '格式例如：包裹文件夹\.minecraft\（最短为空字符串）
        Dim VersionName As String = Match.Groups(1).Value
        Log("[ModPack] 检测到压缩包的 .minecraft 根目录：" & ArchiveBaseFolder & "，命中的版本名：" & VersionName)
        '获取解压路径
        MyMsgBox("接下来请选择一个空文件夹，它会被安装到这个文件夹里。", "安装", "继续", ForceWait:=True)
        Dim TargetFolder As String = SelectFolder("选择安装目标（必须是一个空文件夹）")
        If String.IsNullOrEmpty(TargetFolder) Then Throw New CancelledException
        If TargetFolder.Contains("!") OrElse TargetFolder.Contains(";") Then Hint("Minecraft 文件夹路径中不能含有感叹号或分号！", HintType.Critical) : Throw New CancelledException
        If Directory.GetFileSystemEntries(TargetFolder).Length > 0 Then Hint("请选择一个空文件夹作为安装目标！", HintType.Critical) : Throw New CancelledException
        '解压
        Dim Loader As New LoaderCombo(Of String)("解压压缩包", {
            New LoaderTask(Of String, Integer)("解压压缩包",
            Sub(Task As LoaderTask(Of String, Integer))
                ExtractModpackFiles(TargetFolder, FileAddress, Task, 0.95)
                '加入文件夹列表
                PageSelectLeft.AddFolder(TargetFolder & ArchiveBaseFolder, GetFolderNameFromPath(TargetFolder), False)
                Thread.Sleep(400) '避免文件争用
                RunInUi(Sub() FrmMain.PageChange(FormMain.PageType.VersionSelect))
            End Sub)
        }) With {.OnStateChanged = AddressOf McInstallState}
        Loader.Start(TargetFolder)
        LoaderTaskbarAdd(Loader)
        FrmMain.BtnExtraDownload.ShowRefresh()
        FrmMain.BtnExtraDownload.Ribble()
        Return Loader
    End Function

#End Region

#Region "整合包升级数据迁移"

    ''' <summary>
    ''' 检查并处理同名版本的数据迁移
    ''' </summary>
    ''' <returns>返回处理后的版本名（可能被修改）</returns>
    Private Function CheckAndMigrateOldVersion(VersionName As String) As String
        Dim OldVersionPath As String = $"{PathMcFolder}versions\{VersionName}"

        If Not Directory.Exists(OldVersionPath) Then
            Return VersionName '不存在旧版本，直接返回
        End If

        '检测版本隔离设置
        Dim HasVersionIsolation = Setup.Get("VersionArgumentIndieV2", Version:=New McVersion(OldVersionPath))
        Dim HasImportantData = False

        '快速检查是否有重要数据
        If Directory.Exists(Path.Combine(OldVersionPath, "saves")) AndAlso
           New DirectoryInfo(Path.Combine(OldVersionPath, "saves")).GetDirectories().Any() Then
            HasImportantData = True
        End If

        '根据是否有重要数据调整提示
        Dim Message As String
        If HasImportantData Then
            Message = $"发现已存在同名版本 {VersionName}，检测到该版本包含游戏存档。{vbCrLf}{vbCrLf}是否要迁移旧版本的数据到新版本？这将保留你的存档、设置和mod数据。"
        Else
            Message = $"发现已存在同名版本 {VersionName}。{vbCrLf}{vbCrLf}是否要迁移旧版本的数据到新版本？"
        End If

        Select Case MyMsgBox(Message, "检测到旧版本", "迁移数据并继续", "直接覆盖安装", "取消")
            Case 1 '迁移数据
                '启动迁移流程
                Dim NewVersionName = GetUniqueVersionName(VersionName)
                Dim MigrationLoader = ModpackUpgrade(NewVersionName, OldVersionPath)

                If MigrationLoader IsNot Nothing Then
                    '启动迁移任务
                    MigrationLoader.Start()
                    LoaderTaskbarAdd(MigrationLoader)

                    '等待迁移完成
                    Do While MigrationLoader.State = LoaderState.Loading
                        Thread.Sleep(100)
                    Loop

                    If MigrationLoader.State = LoaderState.Finished Then
                        Log($"[ModPack] 数据迁移成功，新版本名：{NewVersionName}")
                        Return NewVersionName
                    Else
                        Log($"[ModPack] 数据迁移失败，使用原版本名：{VersionName}")
                        If MyMsgBox("数据迁移失败，是否继续覆盖安装？", "迁移失败", "继续", "取消") = 2 Then
                            Throw New CancelledException
                        End If
                    End If
                End If

                Return VersionName

            Case 2 '直接覆盖
                If HasImportantData Then
                    If MyMsgBox($"确定要覆盖现有版本 {VersionName} 吗？{vbCrLf}{vbCrLf}⚠️ 警告：这将永久删除所有存档和数据！",
                              "最终确认", "确定删除", "取消") = 2 Then
                        Throw New CancelledException
                    End If
                End If
                Return VersionName

            Case Else '取消
                Throw New CancelledException
        End Select
    End Function

    ''' <summary>
    ''' 获取一个不重复的版本名
    ''' </summary>
    Private Function GetUniqueVersionName(BaseName As String) As String
        Dim NewName = BaseName & "_new"
        Dim Counter = 1

        While Directory.Exists($"{PathMcFolder}versions\{NewName}")
            Counter += 1
            NewName = BaseName & $"_new{Counter}"
        End While

        Return NewName
    End Function

    '已知的mod数据文件夹映射表
    Private ReadOnly ModDataFolders As New Dictionary(Of String, String) From {
        {"XaeroWaypoints", "Xaero小地图路径点"},
        {"XaeroWorldMap", "Xaero世界地图数据"},
        {"schematics", "Litematica/Schematica建筑蓝图"},
        {"journeymap", "JourneyMap地图数据"},
        {"voxelmap", "VoxelMap地图数据"},
        {"rei", "REI配置"},
        {"jei", "JEI配置"},
        {"emi", "EMI配置"},
        {"worldedit", "WorldEdit数据"},
        {"structurize", "Structurize蓝图"},
        {"create-schematics", "机械动力蓝图"},
        {"blueprints", "BuildCraft蓝图"},
        {"ftbbackups", "FTB备份"},
        {"ftbchunks", "FTB区块数据"},
        {"ftbquests", "FTB任务数据"},
        {"ftb-teams", "FTB团队数据"},
        {"essential", "Essential配置"},
        {"simple-rpc", "Discord RPC配置"},
        {"bobby", "Bobby缓存数据"},
        {"distant-horizons-server-data", "Distant Horizons LOD数据"},
        {"iris", "Iris光影配置"},
        {"meteor-client", "Meteor客户端配置"},
        {"wurst", "Wurst客户端配置"},
        {"baritone", "Baritone路径数据"},
        {"minimap", "小地图数据"},
        {"mapwriter", "MapWriter地图数据"},
        {"betterquesting", "Better Questing任务数据"},
        {"opencomputers", "OpenComputers数据"},
        {"computercraft", "ComputerCraft程序"},
        {"mystcraft", "Mystcraft时代数据"},
        {"thaumcraft", "Thaumcraft研究数据"},
        {"galacticraft", "GalactiCraft数据"},
        {"buildcraft", "BuildCraft配置"},
        {"industrialcraft", "IndustrialCraft配置"},
        {"applied-energistics-2", "AE2配置"},
        {"railcraft", "Railcraft配置"},
        {"tinkers-construct", "匠魂配置"},
        {"botania", "植物魔法配置"},
        {"blood-magic", "血魔法配置"},
        {"astral-sorcery", "星辉魔法配置"},
        {"immersive-engineering", "沉浸工程配置"},
        {"mekanism", "通用机械配置"}
    }

    ''' <summary>
    ''' 整合包升级时的数据迁移入口函数
    ''' </summary>
    Public Function ModpackUpgrade(NewVersionName As String, OldVersionPath As String) As LoaderCombo(Of String)
        Log($"[ModPack] 开始整合包升级数据迁移：从 {OldVersionPath} 到 {NewVersionName}")

        '检测可迁移的数据
        Dim MigrationItems = DetectMigrationItems(OldVersionPath)
        If MigrationItems.Count = 0 Then
            Log("[ModPack] 未检测到可迁移的数据")
            Return Nothing
        End If

        '显示迁移选择对话框
        Dim SelectedItems = ShowMigrationDialog(MigrationItems, OldVersionPath, NewVersionName)
        If SelectedItems Is Nothing OrElse SelectedItems.Count = 0 Then
            Log("[ModPack] 用户取消数据迁移")
            Return Nothing
        End If

        '创建迁移任务
        Dim Loader As New LoaderCombo(Of String)("整合包数据迁移", {
            New LoaderTask(Of String, Integer)("创建备份",
            Sub(Task As LoaderTask(Of String, Integer))
                CreateMigrationBackup(OldVersionPath, SelectedItems, Task)
                Task.Progress = 1
            End Sub) With {.ProgressWeight = 1},
            New LoaderTask(Of String, Integer)("迁移数据",
            Sub(Task As LoaderTask(Of String, Integer))
                MigrateVersionData(OldVersionPath, $"{PathMcFolder}versions\{NewVersionName}", SelectedItems, Task)
                Task.Progress = 1
            End Sub) With {.ProgressWeight = 3}
        })

        Return Loader
    End Function

    ''' <summary>
    ''' 检测可迁移的数据项
    ''' </summary>
    Private Function DetectMigrationItems(VersionPath As String) As List(Of MigrationItem)
        Dim Items As New List(Of MigrationItem)

        '检测标准文件夹
        AddStandardFolderItem(Items, VersionPath, "saves", "游戏存档", True)
        AddStandardFolderItem(Items, VersionPath, "screenshots", "游戏截图", False)
        AddStandardFolderItem(Items, VersionPath, "resourcepacks", "资源包", False)
        AddStandardFolderItem(Items, VersionPath, "shaderpacks", "光影包", False)
        AddStandardFolderItem(Items, VersionPath, "configs", "Mod配置", False)
        AddStandardFolderItem(Items, VersionPath, "config", "Mod配置", False) '两种命名都检查

        '检测标准文件
        AddStandardFileItem(Items, VersionPath, "options.txt", "游戏设置", True)
        AddStandardFileItem(Items, VersionPath, "servers.dat", "服务器列表", True)
        AddStandardFileItem(Items, VersionPath, "hotbar.nbt", "快捷栏", False)

        '检测已知的mod文件夹
        For Each ModFolder In ModDataFolders
            Dim FolderPath = Path.Combine(VersionPath, ModFolder.Key)
            If Directory.Exists(FolderPath) Then
                Dim DirInfo As New DirectoryInfo(FolderPath)
                Items.Add(New MigrationItem With {
                    .RelativePath = ModFolder.Key,
                    .DisplayName = ModFolder.Value,
                    .ItemType = MigrationItemType.ModFolder,
                    .Size = GetDirectorySize(DirInfo),
                    .IsDefault = False,
                    .IsKnown = True
                })
            End If
        Next

        '检测未知但可能的mod文件夹
        Dim RootDir As New DirectoryInfo(VersionPath)
        For Each SubDir In RootDir.GetDirectories()
            If Not IsStandardMinecraftFolder(SubDir.Name) AndAlso
               Not ModDataFolders.ContainsKey(SubDir.Name) Then
                '基于特征判断是否可能是mod数据
                If IsPossibleModDataFolder(SubDir) Then
                    Items.Add(New MigrationItem With {
                        .RelativePath = SubDir.Name,
                        .DisplayName = $"{SubDir.Name} (未识别)",
                        .ItemType = MigrationItemType.UnknownFolder,
                        .Size = GetDirectorySize(SubDir),
                        .IsDefault = False,
                        .IsKnown = False
                    })
                End If
            End If
        Next

        Return Items
    End Function

    ''' <summary>
    ''' 添加标准文件夹项
    ''' </summary>
    Private Sub AddStandardFolderItem(Items As List(Of MigrationItem), BasePath As String,
                                      FolderName As String, DisplayName As String, IsDefault As Boolean)
        Dim FolderPath = Path.Combine(BasePath, FolderName)
        If Directory.Exists(FolderPath) Then
            Dim DirInfo As New DirectoryInfo(FolderPath)
            If DirInfo.GetFiles("*", SearchOption.AllDirectories).Any() Then '有内容才添加
                Items.Add(New MigrationItem With {
                    .RelativePath = FolderName,
                    .DisplayName = DisplayName,
                    .ItemType = MigrationItemType.StandardFolder,
                    .Size = GetDirectorySize(DirInfo),
                    .IsDefault = IsDefault,
                    .IsKnown = True
                })
            End If
        End If
    End Sub

    ''' <summary>
    ''' 添加标准文件项
    ''' </summary>
    Private Sub AddStandardFileItem(Items As List(Of MigrationItem), BasePath As String,
                                    FileName As String, DisplayName As String, IsDefault As Boolean)
        Dim FilePath = Path.Combine(BasePath, FileName)
        If File.Exists(FilePath) Then
            Dim FileInfo As New FileInfo(FilePath)
            Items.Add(New MigrationItem With {
                .RelativePath = FileName,
                .DisplayName = DisplayName,
                .ItemType = MigrationItemType.StandardFile,
                .Size = FileInfo.Length,
                .IsDefault = IsDefault,
                .IsKnown = True
            })
        End If
    End Sub

    ''' <summary>
    ''' 判断是否是标准的Minecraft文件夹
    ''' </summary>
    Private Function IsStandardMinecraftFolder(FolderName As String) As Boolean
        Dim StandardFolders = {"assets", "libraries", "logs", "versions", "crash-reports",
                               "saves", "screenshots", "resourcepacks", "shaderpacks",
                               "mods", "config", "configs", ".mixin.out", "natives",
                               "PCL", "backups", "replay_recordings", "replay_videos"}
        Return StandardFolders.Contains(FolderName.ToLower())
    End Function

    ''' <summary>
    ''' 基于特征判断是否可能是mod数据文件夹
    ''' </summary>
    Private Function IsPossibleModDataFolder(DirInfo As DirectoryInfo) As Boolean
        '跳过空文件夹
        If Not DirInfo.GetFiles("*", SearchOption.AllDirectories).Any() Then Return False

        '跳过过大的文件夹(>2GB)
        If GetDirectorySize(DirInfo) > 2L * 1024 * 1024 * 1024 Then Return False

        '检查文件类型特征
        Dim Files = DirInfo.GetFiles("*", SearchOption.TopDirectoryOnly)
        Dim HasDataFiles = Files.Any(Function(f)
            Return f.Extension.ToLower() In {".dat", ".nbt", ".json", ".yml", ".conf", ".cfg", ".txt", ".schematic"}
        End Function)

        '包含数据文件则可能是mod文件夹
        Return HasDataFiles
    End Function

    ''' <summary>
    ''' 获取文件夹大小
    ''' </summary>
    Private Function GetDirectorySize(DirInfo As DirectoryInfo) As Long
        Try
            Return DirInfo.GetFiles("*", SearchOption.AllDirectories).Sum(Function(f) f.Length)
        Catch ex As Exception
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' 显示迁移选择对话框
    ''' </summary>
    Private Function ShowMigrationDialog(Items As List(Of MigrationItem),
                                        OldPath As String, NewVersion As String) As List(Of MigrationItem)
        '构建对话框内容
        Dim CoreItems = Items.Where(Function(i) i.IsDefault).ToList()
        Dim OptionalItems = Items.Where(Function(i) Not i.IsDefault AndAlso i.ItemType <> MigrationItemType.UnknownFolder).ToList()
        Dim UnknownItems = Items.Where(Function(i) i.ItemType = MigrationItemType.UnknownFolder).ToList()

        Dim Message As New List(Of String)
        Message.Add($"检测到旧版本整合包数据，是否要迁移到新版本 {NewVersion}？")
        Message.Add("")

        If CoreItems.Any() Then
            Message.Add("【核心数据】（推荐迁移）")
            For Each Item In CoreItems
                Message.Add($"• {Item.DisplayName} ({GetFriendlySize(Item.Size)})")
            Next
        End If

        If OptionalItems.Any() Then
            Message.Add("")
            Message.Add("【可选数据】")
            For Each Item In OptionalItems
                Message.Add($"• {Item.DisplayName} ({GetFriendlySize(Item.Size)})")
            Next
        End If

        If UnknownItems.Any() Then
            Message.Add("")
            Message.Add("【未识别文件夹】（谨慎选择）")
            For Each Item In UnknownItems
                Message.Add($"• {Item.DisplayName} ({GetFriendlySize(Item.Size)})")
            Next
        End If

        '显示对话框
        Select Case MyMsgBox(String.Join(vbCrLf, Message), "整合包数据迁移", "全部迁移", "仅迁移核心", "取消")
            Case 1 '全部迁移
                Return Items
            Case 2 '仅核心
                Return CoreItems
            Case Else '取消
                Return Nothing
        End Select
    End Function

    ''' <summary>
    ''' 获取友好的文件大小显示
    ''' </summary>
    Private Function GetFriendlySize(Bytes As Long) As String
        If Bytes < 1024 Then
            Return $"{Bytes} B"
        ElseIf Bytes < 1024 * 1024 Then
            Return $"{Math.Round(Bytes / 1024, 1)} KB"
        ElseIf Bytes < 1024 * 1024 * 1024 Then
            Return $"{Math.Round(Bytes / 1024 / 1024, 1)} MB"
        Else
            Return $"{Math.Round(Bytes / 1024 / 1024 / 1024, 2)} GB"
        End If
    End Function

    ''' <summary>
    ''' 创建迁移备份
    ''' </summary>
    Private Sub CreateMigrationBackup(OldPath As String, Items As List(Of MigrationItem), Task As LoaderTask(Of String, Integer))
        Dim BackupPath = Path.Combine(PathMcFolder, "backups", $"migration_{DateTime.Now:yyyyMMdd_HHmmss}")
        Directory.CreateDirectory(BackupPath)
        Log($"[ModPack] 创建迁移备份：{BackupPath}")

        Dim Progress As Double = 0
        Dim ItemCount = Items.Count

        For Each Item In Items
            Dim SourcePath = Path.Combine(OldPath, Item.RelativePath)
            Dim TargetPath = Path.Combine(BackupPath, Item.RelativePath)

            Try
                If Item.ItemType = MigrationItemType.StandardFile Then
                    CopyFile(SourcePath, TargetPath)
                Else
                    '对于大文件夹，可以考虑使用压缩备份节省空间
                    If Item.Size > 100 * 1024 * 1024 Then '大于100MB
                        Log($"[ModPack] 大文件夹备份：{Item.RelativePath} ({GetFriendlySize(Item.Size)})")
                    End If
                    CopyDirectory(SourcePath, TargetPath)
                End If
            Catch ex As Exception
                Log(ex, $"备份失败：{Item.RelativePath}")
            End Try

            Progress += 1 / ItemCount
            Task.Progress = Progress
        Next
    End Sub

    ''' <summary>
    ''' 执行数据迁移
    ''' </summary>
    Private Sub MigrateVersionData(OldPath As String, NewPath As String,
                                  Items As List(Of MigrationItem), Task As LoaderTask(Of String, Integer))
        Log($"[ModPack] 开始迁移数据：{OldPath} → {NewPath}")

        Dim Progress As Double = 0
        Dim ItemCount = Items.Count

        For Each Item In Items
            Dim SourcePath = Path.Combine(OldPath, Item.RelativePath)
            Dim TargetPath = Path.Combine(NewPath, Item.RelativePath)

            Try
                '检查目标是否已存在
                Dim ShouldCopy = True
                If (File.Exists(TargetPath) OrElse Directory.Exists(TargetPath)) Then
                    '对于某些文件采用合并策略而非覆盖
                    Select Case Item.RelativePath.ToLower()
                        Case "options.txt", "servers.dat"
                            '这些文件询问用户
                            If MyMsgBox($"文件 {Item.DisplayName} 在新版本已存在，是否覆盖？", "文件冲突", "覆盖", "跳过") = 2 Then
                                ShouldCopy = False
                            End If
                        Case "saves"
                            '存档文件夹采用合并策略
                            Log($"[ModPack] 合并存档文件夹")
                            MergeSavesFolder(SourcePath, TargetPath)
                            ShouldCopy = False
                    End Select
                End If

                If ShouldCopy Then
                    If Item.ItemType = MigrationItemType.StandardFile Then
                        CopyFile(SourcePath, TargetPath)
                    Else
                        CopyDirectory(SourcePath, TargetPath, Sub(Delta) Task.Progress += Delta * (1 / ItemCount))
                    End If
                    Log($"[ModPack] 已迁移：{Item.RelativePath}")
                End If
            Catch ex As Exception
                Log(ex, $"迁移失败：{Item.RelativePath}")
            End Try

            Progress += 1 / ItemCount
            If Task.Progress < Progress Then Task.Progress = Progress
        Next

        Log("[ModPack] 数据迁移完成")
        Hint("整合包数据迁移完成！", HintType.Finish)
    End Sub

    ''' <summary>
    ''' 合并存档文件夹
    ''' </summary>
    Private Sub MergeSavesFolder(SourceSaves As String, TargetSaves As String)
        If Not Directory.Exists(TargetSaves) Then
            Directory.CreateDirectory(TargetSaves)
        End If

        For Each SaveDir In New DirectoryInfo(SourceSaves).GetDirectories()
            Dim TargetSaveDir = Path.Combine(TargetSaves, SaveDir.Name)
            If Not Directory.Exists(TargetSaveDir) Then
                '不存在则直接复制
                CopyDirectory(SaveDir.FullName, TargetSaveDir)
            Else
                '存在则重命名
                Dim NewName = SaveDir.Name & "_old"
                Dim Counter = 1
                While Directory.Exists(Path.Combine(TargetSaves, NewName))
                    NewName = SaveDir.Name & $"_old{Counter}"
                    Counter += 1
                End While
                CopyDirectory(SaveDir.FullName, Path.Combine(TargetSaves, NewName))
                Log($"[ModPack] 存档重命名：{SaveDir.Name} → {NewName}")
            End If
        Next
    End Sub

    '数据结构定义
    Private Class MigrationItem
        Public Property RelativePath As String
        Public Property DisplayName As String
        Public Property ItemType As MigrationItemType
        Public Property Size As Long
        Public Property IsDefault As Boolean
        Public Property IsKnown As Boolean
    End Class

    Private Enum MigrationItemType
        StandardFolder
        StandardFile
        ModFolder
        UnknownFolder
    End Enum

#End Region

End Module
