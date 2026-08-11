Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports NXOpen
Imports NXOpen.Assemblies
Imports NXOpen.UF

''' <summary>
''' Read-only assembly preflight for STEP failures caused by missing, unloaded,
''' unavailable, or invalid component occurrences.
'''
''' Target: Siemens NX X 2506, local and Teamcenter-managed sessions.
''' </summary>
Public Module NX_Assembly_Load_Diagnostic

    Private Const ReportFileName As String = "NX_Assembly_Load_Diagnostic_Report.txt"
    Private Const MaxComponentOccurrences As Integer = 100000

    Private Enum DiagnosticPriority
        Ok = 0
        [Error] = 10
        Unloaded = 20
        PrototypeUnavailable = 30
        MissingFile = 40
        InvalidObject = 50
    End Enum

    Private NotInheritable Class ScanNode
        Public Property Component As Component
        Public Property Level As Integer
        Public Property ParentName As String
        Public Property ParentPath As String
    End Class

    Private NotInheritable Class DiagnosticRecord
        Public Property ComponentName As String = "<unavailable>"
        Public Property ParentAssembly As String = "<unavailable>"
        Public Property AssemblyPath As String = "<unavailable>"
        Public Property Level As Integer
        Public Property PartNumber As String = "<not available>"
        Public Property Revision As String = "<not available>"
        Public Property DatasetName As String = "<not available>"
        Public Property ManagedStatus As String = "<not applicable>"
        Public Property FilePath As String = "<not available>"
        Public Property LoadState As String = "<not available>"
        Public Property ReferenceSet As String = "<not available>"
        Public Property IsSuppressed As Boolean
        Public Property Status As String = "OK"
        Public Property Reason As String = "No load-health problem detected."
        Public Property FailedOperation As String = ""
        Public Property ExceptionMessage As String = ""
        Public Property Recommendation As String = "No corrective action required."
        Public ReadOnly Errors As New List(Of String)()

        Private _priority As DiagnosticPriority = DiagnosticPriority.Ok

        Public Sub SetStatus(ByVal statusValue As String,
                             ByVal priorityValue As DiagnosticPriority,
                             ByVal reasonValue As String,
                             ByVal recommendationValue As String)
            If priorityValue < _priority Then Return
            _priority = priorityValue
            Status = statusValue
            Reason = Clean(reasonValue, "No reason was supplied by NX.")
            Recommendation = Clean(recommendationValue, "Review the component in Assembly Navigator.")
        End Sub

        Public Sub AddIssue(ByVal operation As String, ByVal ex As Exception)
            Dim details As String = FormatException(ex)
            Errors.Add(operation & ": " & details)

            If IsInvalidObjectException(ex) Then
                SetStatus(
                    "INVALID_OBJECT",
                    DiagnosticPriority.InvalidObject,
                    "NX rejected an operation on this occurrence or its referenced object.",
                    "Open the reported assembly path, replace or repair the stale occurrence, then reopen and fully load the assembly before STEP export."
                )
                ' An IM0541 signature is more useful than an earlier generic
                ' inspection error, so it becomes the primary failed operation.
                FailedOperation = operation
                ExceptionMessage = details
            Else
                SetStatus(
                    "ERROR",
                    DiagnosticPriority.Error,
                    "NX returned an unexpected error while inspecting this component.",
                    "Review the failed operation and NX exception below; repair the occurrence before STEP export."
                )
                If String.IsNullOrWhiteSpace(FailedOperation) Then
                    FailedOperation = operation
                    ExceptionMessage = details
                End If
            End If
        End Sub
    End Class

    Public Sub Main()
        Dim session As Session = Session.GetSession()
        Dim listingWindow As ListingWindow = session.ListingWindow
        listingWindow.Open()
        listingWindow.WriteLine("NX Assembly Diagnostic Started...")
        listingWindow.WriteLine("")

        Dim workPart As Part = session.Parts.Work
        If workPart Is Nothing Then
            listingWindow.WriteLine("No work part is open. Nothing was scanned.")
            Return
        End If

        Dim assemblyName As String = SafePartName(workPart)
        listingWindow.WriteLine("Scanning assembly:")
        listingWindow.WriteLine(assemblyName)
        listingWindow.WriteLine("")

        Dim records As New List(Of DiagnosticRecord)()
        Dim globalErrors As New List(Of String)()
        Dim isAssembly As Boolean = False

        Try
            isAssembly = ScanAssembly(
                workPart,
                records,
                globalErrors,
                session.IsManagedMode
            )
        Catch ex As Exception
            globalErrors.Add("ScanAssembly: " & FormatException(ex))
        End Try

        Dim failedCount As Integer = 0
        For Each record As DiagnosticRecord In records
            If Not String.Equals(record.Status, "OK", StringComparison.OrdinalIgnoreCase) Then
                failedCount += 1
            End If
        Next

        listingWindow.WriteLine("Components found:")
        listingWindow.WriteLine(records.Count.ToString())
        listingWindow.WriteLine("")
        listingWindow.WriteLine("Errors found:")
        listingWindow.WriteLine((failedCount + globalErrors.Count).ToString())
        listingWindow.WriteLine("")

        If Not isAssembly Then
            listingWindow.WriteLine("The current work part is not an assembly (no child components were found).")
        End If

        Try
            Dim reportPath As String = Path.Combine(ResolveOutputDirectory(), ReportFileName)
            WriteReport(reportPath, workPart, isAssembly, records, globalErrors)
            listingWindow.WriteLine("Report generated successfully.")
            listingWindow.WriteLine(reportPath)
        Catch ex As Exception
            listingWindow.WriteLine("ERROR: The diagnostic report could not be written.")
            listingWindow.WriteLine(FormatException(ex))
        End Try
    End Sub

    Private Function ScanAssembly(ByVal workPart As Part,
                                  ByVal records As List(Of DiagnosticRecord),
                                  ByVal globalErrors As List(Of String),
                                  ByVal sessionIsManaged As Boolean) As Boolean
        Dim root As Component = Nothing
        Try
            If workPart.ComponentAssembly IsNot Nothing Then
                root = workPart.ComponentAssembly.RootComponent
            End If
        Catch ex As Exception
            globalErrors.Add("Get RootComponent: " & FormatException(ex))
            Return False
        End Try

        If root Is Nothing Then Return False

        Dim rootChildren() As Component
        Try
            rootChildren = root.GetChildren()
        Catch ex As Exception
            Dim rootFailure As New DiagnosticRecord With {
                .ComponentName = SafePartName(workPart),
                .ParentAssembly = "<session>",
                .AssemblyPath = SafePartName(workPart),
                .Level = 0
            }
            rootFailure.AddIssue("RootComponent.GetChildren", ex)
            records.Add(rootFailure)
            Return True
        End Try

        If rootChildren Is Nothing OrElse rootChildren.Length = 0 Then Return False

        Dim topName As String = SafePartName(workPart)
        Dim stack As New Stack(Of ScanNode)()
        For index As Integer = rootChildren.Length - 1 To 0 Step -1
            stack.Push(New ScanNode With {
                .Component = rootChildren(index),
                .Level = 1,
                .ParentName = topName,
                .ParentPath = topName
            })
        Next

        While stack.Count > 0
            If records.Count >= MaxComponentOccurrences Then
                globalErrors.Add(
                    "Traversal stopped at the safety limit of " &
                    MaxComponentOccurrences.ToString() & " component occurrences."
                )
                Exit While
            End If

            Dim node As ScanNode = stack.Pop()
            Dim record As DiagnosticRecord = InspectComponent(node, sessionIsManaged)
            Dim children() As Component = Nothing

            Try
                children = node.Component.GetChildren()
            Catch ex As Exception
                record.AddIssue("Component.GetChildren", ex)
            End Try

            records.Add(record)

            If children IsNot Nothing Then
                For index As Integer = children.Length - 1 To 0 Step -1
                    stack.Push(New ScanNode With {
                        .Component = children(index),
                        .Level = node.Level + 1,
                        .ParentName = record.ComponentName,
                        .ParentPath = record.AssemblyPath
                    })
                Next
            End If
        End While

        Return True
    End Function

    Private Function InspectComponent(ByVal node As ScanNode,
                                      ByVal sessionIsManaged As Boolean) As DiagnosticRecord
        Dim record As New DiagnosticRecord With {
            .Level = node.Level,
            .ParentAssembly = Clean(node.ParentName, "<unavailable>")
        }

        Try
            record.ComponentName = Clean(node.Component.DisplayName, "<unnamed component>")
        Catch ex As Exception
            record.AddIssue("Component.DisplayName", ex)
            Try
                record.ComponentName = Clean(node.Component.Name, "<unnamed component>")
            Catch fallbackEx As Exception
                If IsInvalidObjectException(fallbackEx) Then
                    record.AddIssue("Component.Name", fallbackEx)
                End If
            End Try
        End Try

        record.AssemblyPath = AppendPath(node.ParentPath, record.ComponentName)

        ' UF_ASSEM retains the saved child part name even when the higher-level
        ' NXOpen prototype cannot be returned. This fallback is what lets a
        ' local missing file be distinguished from a managed unavailable
        ' prototype without loading or modifying the assembly.
        Dim referencedPartName As String = TryGetReferencedPartName(node.Component, record)
        If Not String.IsNullOrWhiteSpace(referencedPartName) Then
            record.FilePath = referencedPartName
            record.DatasetName = Clean(Path.GetFileName(referencedPartName), referencedPartName)
        End If

        Try
            record.ReferenceSet = Clean(node.Component.ReferenceSet, "<not available>")
        Catch ex As Exception
            record.AddIssue("Component.ReferenceSet", ex)
        End Try

        Try
            record.IsSuppressed = node.Component.IsSuppressed
        Catch ex As Exception
            record.AddIssue("Component.IsSuppressed", ex)
        End Try

        Dim prototype As NXObject = Nothing
        Try
            prototype = node.Component.Prototype
        Catch ex As Exception
            record.AddIssue("Component.Prototype", ex)
        End Try

        If prototype Is Nothing Then
            PopulateIdentityFallbacks(record)
            If sessionIsManaged Then
                record.ManagedStatus = "MANAGED (prototype unavailable)"
            End If

            If String.Equals(record.Status, "INVALID_OBJECT", StringComparison.OrdinalIgnoreCase) Then
                Return record
            End If

            If Not sessionIsManaged AndAlso Not IsUnavailableValue(referencedPartName) Then
                record.SetStatus(
                    "MISSING_FILE",
                    DiagnosticPriority.MissingFile,
                    "The saved local component reference exists in the assembly, but NX cannot resolve a prototype using the current search options.",
                    "Restore the referenced part or correct NX assembly search paths/load options, then reload the assembly."
                )
            Else
                record.SetStatus(
                    "PROTOTYPE_UNAVAILABLE",
                    DiagnosticPriority.PrototypeUnavailable,
                    "The assembly occurrence exists, but NX did not return its prototype.",
                    "Check the Teamcenter revision rule, access rights, dataset availability, or local assembly search paths, then reload the component."
                )
            End If
            Return record
        End If

        Dim prototypePart As BasePart = TryCast(prototype, BasePart)
        If prototypePart Is Nothing Then
            record.SetStatus(
                "PROTOTYPE_UNAVAILABLE",
                DiagnosticPriority.PrototypeUnavailable,
                "The occurrence prototype is not an NX part object.",
                "Replace or repair the occurrence so that it references a valid NX part prototype."
            )
            Return record
        End If

        Try
            Dim prototypeLeaf As String = prototypePart.Leaf
            If Not String.IsNullOrWhiteSpace(prototypeLeaf) Then
                record.DatasetName = prototypeLeaf.Trim()
            End If
        Catch ex As Exception
            record.AddIssue("Prototype.Leaf", ex)
        End Try

        Try
            Dim prototypeFullPath As String = prototypePart.FullPath
            If Not String.IsNullOrWhiteSpace(prototypeFullPath) Then
                record.FilePath = prototypeFullPath.Trim()
            End If
        Catch ex As Exception
            record.AddIssue("Prototype.FullPath", ex)
        End Try

        record.PartNumber = ReadFirstStringAttribute(
            prototypePart,
            New String() {"DB_PART_NO", "ITEM_ID", "ITEM_ID_WITHOUT_REV"},
            record
        )
        record.Revision = ReadFirstStringAttribute(
            prototypePart,
            New String() {"DB_PART_REV", "ITEM_REVISION", "REVISION"},
            record
        )
        Dim releaseStatus As String = ReadFirstStringAttribute(
            prototypePart,
            New String() {"DB_PART_STATUS", "RELEASE_STATUS", "release_status_list"},
            record
        )

        PopulateIdentityFallbacks(record)

        If sessionIsManaged OrElse IsManagedIdentifier(record.FilePath) Then
            record.ManagedStatus = If(
                IsUnavailableValue(releaseStatus),
                "MANAGED (release status not available)",
                releaseStatus
            )
        Else
            record.ManagedStatus = "LOCAL"
        End If

        Try
            record.LoadState = prototypePart.PartLoadState.ToString()
            If prototypePart.PartLoadState = PartLoadState.NotLoaded Then
                record.SetStatus(
                    "UNLOADED",
                    DiagnosticPriority.Unloaded,
                    "The prototype load state is NotLoaded.",
                    "Use Assembly Load Options to fully load this occurrence before STEP export."
                )
            ElseIf prototypePart.PartLoadState = PartLoadState.PartiallyLoaded OrElse
                   prototypePart.PartLoadState = PartLoadState.MinimallyLoaded Then
                record.SetStatus(
                    "UNLOADED",
                    DiagnosticPriority.Unloaded,
                    "The prototype is only " & prototypePart.PartLoadState.ToString() & ".",
                    "Fully load the component with exact geometry before STEP export."
                )
            End If
        Catch ex As Exception
            record.AddIssue("Prototype.PartLoadState", ex)
        End Try

        Try
            If Not prototypePart.IsFullyLoaded Then
                record.SetStatus(
                    "UNLOADED",
                    DiagnosticPriority.Unloaded,
                    "NX reports that the prototype is not fully loaded.",
                    "Fully load the component and its required children before STEP export."
                )
            End If
        Catch ex As Exception
            record.AddIssue("Prototype.IsFullyLoaded", ex)
        End Try

        If IsLocalPartPath(record.FilePath) AndAlso
           Path.IsPathRooted(record.FilePath) AndAlso
           Not File.Exists(record.FilePath) Then
            record.SetStatus(
                "MISSING_FILE",
                DiagnosticPriority.MissingFile,
                "The referenced local part file cannot be found at the recorded path.",
                "Restore the file or correct NX assembly search paths/load options, then reload the assembly."
            )
        End If

        If record.IsSuppressed AndAlso String.Equals(record.Status, "OK", StringComparison.OrdinalIgnoreCase) Then
            record.Reason = "Suppressed occurrence; it is excluded from the active assembly state."
        End If

        Return record
    End Function

    Private Function TryGetReferencedPartName(ByVal component As Component,
                                              ByVal record As DiagnosticRecord) As String
        Try
            Dim ufSession As UFSession = UFSession.GetUFSession()
            Dim instanceTag As Tag = ufSession.Assem.AskInstOfPartOcc(component.Tag)
            Dim partFileSpec As String = ""
            ufSession.Assem.AskPartNameOfChild(instanceTag, partFileSpec)
            Return Clean(partFileSpec, "")
        Catch ex As Exception
            If IsInvalidObjectException(ex) Then
                record.AddIssue("UFAssem.AskPartNameOfChild", ex)
            End If
            Return ""
        End Try
    End Function

    Private Function ReadFirstStringAttribute(ByVal target As NXObject,
                                              ByVal titles() As String,
                                              ByVal record As DiagnosticRecord) As String
        For Each title As String In titles
            Try
                Dim value As String = target.GetStringAttribute(title)
                If Not String.IsNullOrWhiteSpace(value) Then Return value.Trim()
            Catch ex As Exception
                ' A missing optional attribute normally raises in NX. Only retain the
                ' exception when it is the OM-object failure this journal diagnoses.
                If IsInvalidObjectException(ex) Then
                    record.AddIssue("Prototype.GetStringAttribute(" & title & ")", ex)
                    Exit For
                End If
            End Try
        Next
        Return "<not available>"
    End Function

    Private Sub PopulateIdentityFallbacks(ByVal record As DiagnosticRecord)
        If IsManagedIdentifier(record.FilePath) Then
            Dim normalized As String = record.FilePath.Replace("\"c, "/"c)
            Dim marker As Integer = normalized.IndexOf("@DB/", StringComparison.OrdinalIgnoreCase)
            If marker >= 0 Then
                Dim managedPath As String = normalized.Substring(marker + 4)
                Dim tokens() As String = managedPath.Split("/"c)
                If IsUnavailableValue(record.PartNumber) AndAlso tokens.Length > 0 Then
                    record.PartNumber = Clean(tokens(0), "<not available>")
                End If
                If IsUnavailableValue(record.Revision) AndAlso tokens.Length > 1 Then
                    record.Revision = Clean(tokens(1), "<not available>")
                End If
            End If
        ElseIf IsUnavailableValue(record.PartNumber) AndAlso Not IsUnavailableValue(record.DatasetName) Then
            record.PartNumber = Path.GetFileNameWithoutExtension(record.DatasetName)
        End If
    End Sub

    Private Function ResolveOutputDirectory() As String
        Dim configured As String = Environment.GetEnvironmentVariable("NX_JOURNALS_IO_DIR")
        Dim outputDirectory As String = configured

        If String.IsNullOrWhiteSpace(outputDirectory) Then
            outputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        End If
        If String.IsNullOrWhiteSpace(outputDirectory) Then
            outputDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        End If
        If String.IsNullOrWhiteSpace(outputDirectory) Then
            Throw New InvalidOperationException("No writable output directory could be resolved.")
        End If

        Directory.CreateDirectory(outputDirectory)
        Return outputDirectory
    End Function

    Private Sub WriteReport(ByVal reportPath As String,
                            ByVal workPart As Part,
                            ByVal isAssembly As Boolean,
                            ByVal records As List(Of DiagnosticRecord),
                            ByVal globalErrors As List(Of String))
        Dim failedCount As Integer = 0
        Dim statusCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

        For Each record As DiagnosticRecord In records
            If Not statusCounts.ContainsKey(record.Status) Then statusCounts(record.Status) = 0
            statusCounts(record.Status) += 1
            If Not String.Equals(record.Status, "OK", StringComparison.OrdinalIgnoreCase) Then
                failedCount += 1
            End If
        Next

        Using writer As New StreamWriter(reportPath, False, New UTF8Encoding(True))
            writer.WriteLine("================================================")
            writer.WriteLine("NX Assembly Load Diagnostic Report")
            writer.WriteLine("================================================")
            writer.WriteLine("")
            writer.WriteLine("Assembly:")
            writer.WriteLine(SafePartName(workPart))
            writer.WriteLine("")
            writer.WriteLine("Generated:")
            writer.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss zzz"))
            writer.WriteLine("")
            writer.WriteLine("Assembly detected:")
            writer.WriteLine(If(isAssembly, "YES", "NO"))
            writer.WriteLine("")
            writer.WriteLine("Total components scanned:")
            writer.WriteLine(records.Count.ToString())
            writer.WriteLine("")
            writer.WriteLine("Failed components:")
            writer.WriteLine(failedCount.ToString())
            writer.WriteLine("")
            writer.WriteLine("Global scan errors:")
            writer.WriteLine(globalErrors.Count.ToString())
            writer.WriteLine("")

            writer.WriteLine("Status summary:")
            Dim orderedStatuses() As String = {
                "OK", "MISSING_FILE", "PROTOTYPE_UNAVAILABLE", "UNLOADED", "INVALID_OBJECT", "ERROR"
            }
            For Each statusValue As String In orderedStatuses
                Dim count As Integer = 0
                If statusCounts.ContainsKey(statusValue) Then count = statusCounts(statusValue)
                writer.WriteLine(statusValue & ": " & count.ToString())
            Next
            writer.WriteLine("")

            If Not isAssembly Then
                writer.WriteLine("Result:")
                writer.WriteLine("The current work part is not an assembly; no child components were found.")
                writer.WriteLine("")
            End If

            If globalErrors.Count > 0 Then
                writer.WriteLine("------------------------------------------------")
                writer.WriteLine("GLOBAL SCAN ERRORS")
                For Each scanError As String In globalErrors
                    writer.WriteLine("- " & scanError)
                Next
                writer.WriteLine("")
            End If

            writer.WriteLine("================================================")
            writer.WriteLine("COMPONENT DETAILS")
            writer.WriteLine("================================================")

            If records.Count = 0 Then
                writer.WriteLine("No component occurrences were available to report.")
            End If

            For Each record As DiagnosticRecord In records
                writer.WriteLine("")
                writer.WriteLine("------------------------------------------------")
                WriteField(writer, "Component", record.ComponentName)
                WriteField(writer, "Parent assembly", record.ParentAssembly)
                WriteField(writer, "Assembly path", record.AssemblyPath)
                WriteField(writer, "Level", record.Level.ToString())
                WriteField(writer, "Part number / Item ID", record.PartNumber)
                WriteField(writer, "Revision", record.Revision)
                WriteField(writer, "Dataset / prototype", record.DatasetName)
                WriteField(writer, "File path / managed ID", record.FilePath)
                WriteField(writer, "Managed status", record.ManagedStatus)
                WriteField(writer, "Load state", record.LoadState)
                WriteField(writer, "Reference set", record.ReferenceSet)
                WriteField(writer, "Suppressed", If(record.IsSuppressed, "YES", "NO"))
                WriteField(writer, "Status", record.Status)
                WriteField(writer, "Reason", record.Reason)
                WriteField(writer, "Recommended corrective action", record.Recommendation)

                If Not String.IsNullOrWhiteSpace(record.FailedOperation) Then
                    WriteField(writer, "Failed operation", record.FailedOperation)
                    WriteField(writer, "Exception", record.ExceptionMessage)
                End If

                If record.Errors.Count > 1 Then
                    writer.WriteLine("Additional inspection errors:")
                    For Each componentError As String In record.Errors
                        writer.WriteLine("- " & componentError)
                    Next
                End If
            Next

            writer.WriteLine("")
            writer.WriteLine("================================================")
            writer.WriteLine("End of report")
            writer.WriteLine("================================================")
        End Using
    End Sub

    Private Sub WriteField(ByVal writer As TextWriter,
                           ByVal label As String,
                           ByVal value As String)
        writer.WriteLine(label & ":")
        writer.WriteLine(Clean(value, "<not available>"))
        writer.WriteLine("")
    End Sub

    Private Function IsInvalidObjectException(ByVal ex As Exception) As Boolean
        Dim message As String = ex.Message
        If message Is Nothing Then message = ""
        Return message.IndexOf("IM0541", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               message.IndexOf("invalid or unsuitable OM object", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               (message.IndexOf("invalid", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso
                message.IndexOf("OM object", StringComparison.OrdinalIgnoreCase) >= 0)
    End Function

    Private Function FormatException(ByVal ex As Exception) As String
        Dim message As String = Clean(ex.Message, ex.GetType().FullName)
        Dim nxError As NXException = TryCast(ex, NXException)
        If nxError IsNot Nothing Then
            Return "NXException " & nxError.ErrorCode.ToString() & ": " & message
        End If
        Return ex.GetType().FullName & ": " & message
    End Function

    Private Function SafePartName(ByVal part As BasePart) As String
        Try
            If Not String.IsNullOrWhiteSpace(part.Leaf) Then Return part.Leaf.Trim()
        Catch
        End Try
        Try
            If Not String.IsNullOrWhiteSpace(part.Name) Then Return part.Name.Trim()
        Catch
        End Try
        Return "<unavailable work part>"
    End Function

    Private Function AppendPath(ByVal parentPath As String, ByVal childName As String) As String
        Dim parentValue As String = Clean(parentPath, "<root>")
        Dim childValue As String = Clean(childName, "<unnamed component>")
        Return parentValue & " / " & childValue
    End Function

    Private Function IsManagedIdentifier(ByVal value As String) As Boolean
        If String.IsNullOrWhiteSpace(value) Then Return False
        Dim normalized As String = value.Replace("\"c, "/"c)
        Return normalized.IndexOf("@DB/", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    Private Function IsLocalPartPath(ByVal value As String) As Boolean
        If IsUnavailableValue(value) OrElse IsManagedIdentifier(value) Then Return False
        Return value.EndsWith(".prt", StringComparison.OrdinalIgnoreCase) OrElse Path.IsPathRooted(value)
    End Function

    Private Function IsUnavailableValue(ByVal value As String) As Boolean
        Return String.IsNullOrWhiteSpace(value) OrElse value.StartsWith("<", StringComparison.Ordinal)
    End Function

    Private Function Clean(ByVal value As String, ByVal fallback As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return fallback
        Return value.Trim()
    End Function

    Public Function GetUnloadOption(ByVal dummy As String) As Integer
        Return CInt(Session.LibraryUnloadOption.Immediately)
    End Function

End Module
