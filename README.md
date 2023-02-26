<b>GAM Command Manager</b>

A Powershell script to provide a more interactive interface of some of the functions of GAM (Google Apps Manager)

This script requires GAM (Google Apps Manager), find GAM here https://github.com/GAM-team/GAM

<b>What is this?</b>

This script provides a menu-based selection of common options under GAM to allow a workspace administrator to manage their Google Workspace Domain without having to type in commands.

<b>How to use it?</b>

Run the PowerShell script either locally on the server (recommended) or through a PS-Session. The script contains options at the top of the file, for example, GAM location that can be hard coded. If you do not save these, you will be prompted to input these options each time you open the script.

Once you've opened the script you will be greeted with a selection menu, currently, GAM Command Manager supports 3 'modules', Google Classroom, Gmail and Calendar. Each module has various commands available. All output of commands is from GAM directly.

![image](https://user-images.githubusercontent.com/89160322/221432910-7aa0647d-1baa-4298-b3bf-5951ee3f08af.png)

if you get the following warning 'GAM Command Manager.ps1 cannot be loaded because running scripts is disabled
on this system.'. Please open up an administrative PowerShell prompt and enter: 'Set-ExecutionPolicy RemoteSigned'. Please be aware of the security implementations of doing so.

---

Items marked [WIP] are currently not implemented. If you come across any bugs or feature requests, please submit an issue.
