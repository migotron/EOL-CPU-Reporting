# EOL-CPU-Reporting
Project i did for work

The point of this project is to automate the process of the EOL reports done for customers.

The process is currently export the data report from Automate to an excel file.
Convert the data into a table.
Review the device's hardware to determine if replacement or upgrades are needed.
If the CPU is older than 5 years, older than 10th gen Intel, or older than 4th gen AMD ryzen, then mark PC as EOL in highlighted Red the entire table row.
Otherwise we move to servers to mark them as servers, we check the Agent Type and highlight Blue for servers the entire table row
Once we marked the EOL PCs and servers, we can focus on the hardware upgrades such as RAM and Storage, which are marked only in the cell.
If the PC has less than 16GBs of RAM in Agent Memory Total, then we highlight the cell Purple for RAM upgrade.
If the PC has less than 25% C Drive Free Space/Percent, then we highlight the cell Light Blue for SSD upgrade.
If the PC does not have pro version of Windows ie Windows 10 Home or Windows 11 Home, then we highlight the cell Orange for Needs Pro version
If the PC does not have any issues like above and it is on Windows 10 Pro, then we can highlight the entire table row as Yellow for "Can be upgraded to Windows 11 Pro", except for the already highlighted cells, which we leave with their already fill cell color.
If the PC does not have any issues like above and it is on Windows 11 Pro already, then we can highlight the entire table row as Green for "Already on windows 11 Pro", except for the already highlighted cells, which we leave with their already fill cell color.

Table columns are: 
Client  
Location	 
Computer Name	
User	
Agent Type	
Agent IP Address	
Manufacturer	
Agent Mainboard	
Agent OS	
Agent Memory Total	
Agent Serial Number	
CPU	
C Drive Total Space	
C Drive Free Space	
C Drive Free Percent	
Total Internal Drive
