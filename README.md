# OASG
Outlook Automatic Signature Generator


The problem:
So far I searched a lot about using Outlook signatures into a corporate way without finding something that satisfy me. There is some tricks but they are not smooth or proffessional, it also have some cons and there is no embed solution from Microsoft Exchange. The best solution is to buy a software, they are good but for sure not free and I was a little afraid to pay from something so simple.


The solution:
This Powershell script is literraly taking informations from Active Directory users, and remplace HTML content with these informations. After that it will put the signature into Outlook and make it the default one.
It is intended to be runned on user side because it take information from the account where it run, it can be run manualy, semi or fully automatised.


Here is how I made it worked for me:
1) We have and OU where all our users reside, I made delegate access to our HR department so they can fill the function, the phone, the location etc... (it require acces to the console, I made a simple RDP connection that launch it for the users). I didn't tryed but maybe you can do it into Outlook if you set up the rights or maybe even from OWA ECP.
2) I made a network share, I limited access to the Domain users groups to avoid unexpected result but you can do it to whatever group you want. In this share I've an html file with the signature template, they also the PowerShell script and the .bat file.
3) I put the .bat script into a GPO with our startup scripts, keep care to put into the user env. and not the computer env.

Point 3) is optional, you can just send the link to the script to the users or on a intranet, so they can update they sign manually whenever they want. Puting it like I did just ensure the signs are keeping up to date with current users informations and with new templates (just replace the html file !)

The batch file just allow you to execute the PowerShell script without headhaches.

The HTML file is a simple signature, nothing special, the trick is to select the right parts with the script.


Feel free to improve it !
