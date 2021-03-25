# :computer: network_gathering_tool_v2 (NGTv2)
The Improved version of Network Gathering Tool (NGTv1) for https://github.com/Dino-s26/Network_Gathering_tool

:blue_book: This new improved version has new feature such as :
1. ++ Exception for unhandled error in previous NGTv1, that make the script stop / not working if unhandled error occured.
2. ++ Improved clean code, better readability and separated function for better troubleshooting.
3. ++ As the previous NGTv1 using hard coded source (separated source file for username, hostname, command, etc) that are time consuming to update related script, now introducing local database using excel file, so no need to recompile the code (unless you want to change the output filename), you just need to edit the excel file (template file provided) if there is any changes within your environment (ip, username, command, etc).
4. ++ Generate log file when exception are happen to make it easy to troubleshoot.
5. **currently still improving this version, and adding more feature to make it easier for Network Engineer usage.*

:clipboard: How to use ?
1. Install related library before compile into .exe or run it as .py. 
The related library as follows :
- paramiko
- openpyxl
- pathlib / pathlib2
- pyinstaller (if you need to compile into .exe file)

2. Update the template file to your needs :
- Filename for the excel source file (filename section)
- Device purpose (create_log function)
- Download the excel file template and update it to your need and save the file into name you want and update it on filename section of the .py file, aslo put the file within the .py script for convinience.

3. Run the .py file :) 
4. (optional) For .exe compile you can use "pyinstaller -F --clean <.py file>", this will compile the script into .exe file and later can be use for task scheduler in windows (windows only) or if you don't want to run it as .py file :)

That's all, I hope this new version could ease our job as Network Engineer and cheers 🍻.
Let me know if you facing issue, maybe I could help to resolved it 😊.

-- End of Readme.md --
There's nothing after this line 😲
