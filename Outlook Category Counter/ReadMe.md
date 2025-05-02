# **Automating Category Counts in Outlook – Setup Guide**
Zander O'Callaghan - 02/05/2025

## **Overview**
This VBA macro will automatically count the number of emails in your Outlook inbox by category, organized by month. Instead of manually tracking these counts, you can run the macro in seconds—saving time and effort. This was designed to assist the AfterCare team in their monthly statistics gathering. Previously, categories in Outlook had to be manually counted to contribute to understanding the team's workload.

## **Benefits**
✅ Automatically organizes email counts by month  
✅ Works with either predefined categories or dynamically finds them  
✅ Filters emails to only look at the last six weeks for efficiency  
✅ Includes error handling for smoother debugging  
✅ Eliminates the need for manual tracking  

## **Setup Steps**

### **1. Open the VBA Editor**
1. Open **Microsoft Outlook**.
2. Press **ALT + F11** to open the VBA editor.
3. Click **Insert > Module** to add a new module.

### **2. Paste the Macro Code**
Copy and paste the VBA macro into the new module:

---

### **3. Run the Macro**
1. In the VBA editor, press **ALT + F8**.
2. Select `CountCategoriesByMonth` from the macro list.
3. Click **Run**.
4. A pop-up will display the results, showing email counts per category by month.

### **Alternative Steps to Access the Developer Tab**
If the Developer tab is not visible in your Outlook ribbon, follow these steps to enable it:

1. Open **Microsoft Outlook**.
2. Click on **File > Options**.
3. In the Outlook Options window, select **Customize Ribbon** from the left-hand menu.
4. On the right side, under **Customize the Ribbon**, check the box for **Developer**.
5. Click **OK** to save your changes.

Once the Developer tab is enabled, you can run the marco by clicking on **Developer > Macros** and selecting it from the list. 

---

## **Customization Options**
- **Dynamic vs. Predefined Categories**:  
  - To manually set expected categories, change `usePredefinedCategories = True` and update the `predefinedCategories` array.
  - To automatically detect categories in your mailbox, leave it as `False`.

- **Change the Folder**:  
  - If you want to analyze emails in a different folder (e.g., Sent Items), change:
    ```vba
    Set olFolder = olNamespace.GetDefaultFolder(olFolderInbox)
    ```
    To:
    ```vba
    Set olFolder = olNamespace.GetDefaultFolder(olFolderSentMail)
    ```

- **Dynamic Folder Selection**:  
  - To allow the user to select a folder dynamically when running the macro, set:
    ```vba
    useDynamicFolderSelection = True
    ```
  - To use a predefined folder (e.g., "Nuaire AfterSales"), set:
    ```vba
    useDynamicFolderSelection = False
    ```

- **Export Data Instead of Showing a Pop-Up**:  
  - Instead of `MsgBox`, modify the macro to write results to an **Excel sheet** or a **text file** for easier record-keeping.

---

## **Error Handling**
- The macro includes error handling to ensure smooth execution:
  - If the maximum email count (`maxEmails`, default: 3000) is exceeded, the macro will stop and display the message:  
    **"Max email count has been reached, please contact IT."**
  - For other errors, a message box will display the error description to help with troubleshooting.

---

## **Limitations**
- The macro is designed to process up to 3000 emails (`maxEmails`) for performance reasons. If your mailbox contains more emails, consider increasing this limit or filtering emails further.
- The macro only processes emails from the last six weeks to improve efficiency.

---
### **Why Use This?**
✅ **Saves time (reportedly about an hour each month!)**  
✅ **Reduces errors**  
✅ **Easy to run anytime**   
✅ **Far less soul destroying**  

If you have any questions or want to tweak the macro further, just let me know! 😊