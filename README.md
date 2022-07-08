# ThankYouFolderOfShame

Microsoft Outlook VB program to dump reply all "Thank you" messages from coworkers into a separate folder

1. Open regedit

1a. Navigate to Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Security

1b. Right click, new -> DWORD and add EnableUnsafeClientMailRules

1c. Right click EnableUnsafeClientMailRules and change value to 1

 
2. In Outlook create a new folder in your inbox titled "Thank you" (without quotes)

 
3. Create a sub routine in Outlook Visual Basic

3a. Enable Developer Mode: File -> tab -> Options -> Customize Ribbon -> Developer check box

3b. Go to developer tab and click Visual Basic

3c. Under the Project window on the left, right click -> insert -> module

3d. Copy all code from TyMacro into new module you created. Save project and exit Visual Basic

 
4. Create your new rule

4a. Go to File -> Manage Rules and Alerts

4b. Add a new rule called Thank you and set Rule description to:

"Apply this rule after the message arrives

on this computer only

run _script_ (click on underscored script in Rules description section and select the script you saved YourProject.CustomMailMessageRule)

(note: Visual Basic window might need to be open to select script)

except if sent only to me (you can exclude this if you want to not see any thank you emails whatsoever)"

 
5. Enjoy!
