# EOL-CPU-Reporting
Project i did for work

The point of this project is to automate the process of the EOL reports done for customers.

The process is currently export the data report from Automate to an excel file.
Convert the data into a table.
Review the device's hardware to determine if replacement or upgrades are needed.
If the CPU is older than 5 years, older than 10th gen Intel, or older than 4th gen AMD ryzen, then mark PC as EOL in highlighted Red the entire table row.


What You Need to Do
1. Update the file path in this line:

     `Set eolWB = Workbooks.Open("C:\Path\To\EOL_CPU_List.xlsx")`
   
   Replace it with the actual path to your EOL CPU list file.
3. Ensure the EOL list is in column A of Sheet1 in that file.
4. Run the macro:
     Press Alt + F11 to open the VBA editor.
     Insert a new module (Insert > Module).
     Paste the code.
     Press F5 or run it from Excel via Alt + F8.


	


	
	
		TODO LIST: 	
		Otherwise we move to servers to mark them as servers, we check the Agent Type and highlight Blue for servers the entire table row

		Once we marked the EOL PCs and servers, we can focus on the hardware upgrades such as RAM and Storage, which are marked only in the cell.

		If the PC has less than 16GBs of RAM in Agent Memory Total, then we highlight the cell Purple for RAM upgrade.

		If the PC has less than 25% C Drive Free Space/Percent, then we highlight the cell Light Blue for SSD upgrade.

		If the PC does not have pro version of Windows ie Windows 10 Home or Windows 11 Home, then we highlight the cell Orange for Needs Pro version

		If the PC does not have any issues like above and it is on Windows 10 Pro, then we can highlight the entire table row as Yellow for "Can be upgraded to Windows 11 Pro", except for the already highlighted cells, which we leave with their already fill cell color.

		If the PC does not have any issues like above and it is on Windows 11 Pro already, then we can highlight the entire table row as Green for "Already on windows 11 Pro", except for the already highlighted cells, which we leave with their already fill cell color.
