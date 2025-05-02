Sub CountCategoriesByMonth()
    On Error GoTo ErrorHandler ' Add error handling

    Dim olNamespace As Outlook.NameSpace
    Dim olFolder As Outlook.Folder
    Dim olItems As Outlook.Items
    Dim restrictedItems As Outlook.Items
    Dim olItem As Object
    Dim categoryCounts As Object
    Dim itemDate As Date
    Dim monthYearKey As Variant
    Dim category As Variant
    Dim output As String
    Dim sixWeeksAgo As Date
    Dim predefinedMonths As Variant
    Dim predefinedCategories As Variant
    Dim usePredefinedCategories As Boolean
    Dim useDynamicFolderSelection As Boolean
    Dim i As Long
    Dim emailCount As Long
    Const maxEmails As Long = 3000

    ' Toggle between predefined and dynamic categories. True for predefined, False for dynamic
    usePredefinedCategories = False

    ' Toggle between dynamic folder selection and predefined folder
    useDynamicFolderSelection = False ' Set to True to enable dynamic folder selection

    ' Initialize variables
    Set olNamespace = Application.GetNamespace("MAPI")

    If useDynamicFolderSelection Then
        ' Allow user to select folder dynamically
        Set olFolder = Application.Session.PickFolder
        If olFolder Is Nothing Then Exit Sub ' User canceled folder selection
    Else
        ' Use predefined folder
        Set olFolder = olNamespace.GetDefaultFolder(olFolderInbox)
        'Set olFolder = olNamespace.Folders("Nuaire AfterSales").Folders("Inbox")
    End If

    Set olItems = olFolder.Items
    sixWeeksAgo = DateAdd("ww", -6, Date) ' Calculate the date six weeks ago

    ' Restrict the items to the last six weeks
    olItems.Sort "[ReceivedTime]", True ' Sort items by ReceivedTime in descending order
    Set restrictedItems = olItems.Restrict("[ReceivedTime] >= '" & Format(sixWeeksAgo, "yyyy-mm-dd") & "'")

    ' Create the dictionary to hold category counts
    Set categoryCounts = CreateObject("Scripting.Dictionary")

    ' Predefine the months within the 6-week range
    predefinedMonths = Array(Format(DateAdd("ww", -6, Date), "mm-yyyy"), _
                             Format(DateAdd("ww", -6, DateAdd("m", 1, Date)), "mm-yyyy"), _
                             Format(Date, "mm-yyyy"))

    predefinedCategories = Array("Category1", "Category2", "Category3")

    If usePredefinedCategories Then
        For Each category In predefinedCategories
            If Not categoryCounts.exists(category) Then
                Set categoryCounts(category) = CreateObject("Scripting.Dictionary")
            End If
            For i = LBound(predefinedMonths) To UBound(predefinedMonths)
                monthYearKey = predefinedMonths(i)
                If Not categoryCounts(category).exists(monthYearKey) Then
                    categoryCounts(category)(monthYearKey) = 0
                End If
            Next i
        Next category
    End If

    ' Iterate through restricted items
    emailCount = 0

    For Each olItem In restrictedItems
        If emailCount >= maxEmails Then
            Err.Raise vbObjectError + 1, "CountCategoriesByMonth", "Max email count has been reached, please contact IT"
        End If

        If TypeOf olItem Is Outlook.MailItem Then
            itemDate = olItem.ReceivedTime
            monthYearKey = Format(itemDate, "mm-yyyy") ' Change format to MM-yyyy

            ' Split categories (items can have multiple categories)
            If olItem.Categories <> "" Then
                For Each category In Split(olItem.Categories, ",")
                    category = Trim(category) ' Clean up spacing
                    If Not categoryCounts.exists(category) Then
                        ' Dynamically add categories if not using predefined ones
                        If Not usePredefinedCategories Then
                            Set categoryCounts(category) = CreateObject("Scripting.Dictionary")
                        Else
                            GoTo SkipCategory
                        End If
                    End If

                    ' Count items by month
                    If Not categoryCounts(category).exists(monthYearKey) Then
                        categoryCounts(category)(monthYearKey) = 0
                    End If
                    categoryCounts(category)(monthYearKey) = categoryCounts(category)(monthYearKey) + 1
SkipCategory:
                Next category
            End If
        End If

        emailCount = emailCount + 1
    Next olItem

    ' Generate output
    For Each category In categoryCounts.Keys
        output = output & "Category: " & category & vbCrLf
        For Each monthYearKey In categoryCounts(category).Keys
            output = output & "  " & monthYearKey & ": " & categoryCounts(category)(monthYearKey) & " items" & vbCrLf
        Next monthYearKey
        output = output & vbCrLf
    Next category

    ' Display results
    MsgBox output, vbInformation, "Category Counts by Month"

    Exit Sub

ErrorHandler:
    If Err.Number = vbObjectError + 1 Then
        MsgBox Err.Description, vbCritical, "Error"
    Else
        MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    End If
End Sub