**Week-14**
Create a UiPath workflow to remind the user to take their medicine every 4 hours using the following activities in order: 
Comment Out, Message Box, Assign, Comment Out, Send Gmail, Delay, Message Box and use every activity
![image](https://github.com/user-attachments/assets/2ec1ed7f-0b15-42dd-9a00-0a223305e042)

**Week-13**
Design a process to Generate Covid-19 report and send this report to the required recipient using the following activites in order:
Generate Data Table From Text, Output Data Table as Text, Message Box, For Each Row in Data Table, ( Get Row Item, Get Row Item, Get Row Item, Get Row Item, Get Row Item, Get Row Item,) -> column number 0 to 4 ( region, testresult, testtype, name, emailid),  Browser(path ms word), Word AppGcation Scope, Assign, Append Text, Save Document as PDF, Assign, Send Google mail and use every activity one by one with out any error or warning 
**Body** "Hi " + name + "," + Environment.NewLine + "Please find the lab report attached." + Environment.NewLine + "Thanks"

**Week-12**
Create a UiPath workflow using Modern Excel activities to:
Use the Use Excel File activity:
File name: "medicinexl.xlsx"
Reference name: Excel
Options:
✅ Read formatting
✅ Save changes
✅ Create if not exists
Inside the Use Excel File, add the Set Range Color activity:
Sheet Name: "Sheet1"
Range: "A1:D2"
Color: System.Drawing.Color.Orange
Save and run the workflow to confirm that the background color of the specified range is updated successfully.
![image](https://github.com/user-attachments/assets/1664e835-d64b-418b-88bc-b0227d52f4c0)

**Week-11**
Design a process to read all PDF files from a folder and then close them all:
Activities:
For Each File in Folder → Message Box → Extract PDF → Write Text File → Write Text File → Message Box 
![image](https://github.com/user-attachments/assets/ae2837f4-3a4a-4320-9379-b65594f2d0b3)



