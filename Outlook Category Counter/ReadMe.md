# **Automating Category Counts in Outlook – Setup Guide**
ZanderO'Callaghan - 02/05/2025

## **Overview**
This VBA macro will automatically count the number of emails in your Outlook inbox by category, organized by month. Instead of manually tracking these counts, you can run the macro in seconds—saving time and effort. This was designed to assist the AfterCare team in their monthly statistics gathering. Previsouly categories in Outlook had to be manually counted to contribute to understanding the teams workload.

## **Benefits**
✅ Automatically organizes email counts by month  
✅ Works with either predefined categories or dynamically finds them  
✅ Filters emails to only look at the last six weeks for efficiency  
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

---

### **Customization Options**
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

- **Export Data Instead of Showing a Pop-Up**:  
  - Instead of `MsgBox`, modify the macro to write results to an **Excel sheet** or a **text file** for easier record-keeping.

---

### **Why Use This?**
✅ **Eliminates manual category counting**  
✅ **Saves time (reportedly about an hour each month!)**  
✅ **Reduces errors**  
✅ **Easy to run anytime**

If you have any questions or want to tweak the macro further, just let me know!😊