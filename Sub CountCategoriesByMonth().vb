Sub CountCategoriesByMonth()
    Dim olNamespace As Outlook.NameSpace
    Dim olFolder As Outlook.Folder
    Dim olItems As Outlook.Items
    Dim olItem As Object
    Dim categoryCounts As Object
    Dim itemDate As Date
    Dim monthYearKey As Variant
    Dim category As Variant
    Dim output As String
    
    ' Initialize variables
    Set olNamespace = Application.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(olFolderInbox) ' Change to desired folder
    Set olItems = olFolder.Items
    Set categoryCounts = CreateObject("Scripting.Dictionary")
    
    ' Iterate through each item in the folder
    For Each olItem In olItems
        If TypeOf olItem Is Outlook.MailItem Then
            itemDate = olItem.ReceivedTime
            monthYearKey = Format(itemDate, "yyyy-mm")
            
            ' Split categories (items can have multiple categories)
            If olItem.Categories <> "" Then
                For Each category In Split(olItem.Categories, ",")
                    category = Trim(category) ' Clean up spacing
                    If Not categoryCounts.exists(category) Then
                        Set categoryCounts(category) = CreateObject("Scripting.Dictionary")
                    End If
                    
                    ' Count items by month
                    If Not categoryCounts(category).exists(monthYearKey) Then
                        categoryCounts(category)(monthYearKey) = 0
                    End If
                    categoryCounts(category)(monthYearKey) = categoryCounts(category)(monthYearKey) + 1
                Next category
            End If
        End If
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
End Sub
