Microsoft Outlook on Mac does not allow creation of server or client rules that permanently block email from specific sender from accessing your mailbox.
It is a problem if you get thousands automated emails per day.
The solution is following:
- create server rule that delete emails from specific sender(s) bu moving them into Trash folder
- run
      osacompile -o PurgeSpam.app PurgeSpam.scpt 
  to compile PurgeSpam.scpt  and to generate PurgeSpam.app
  
At this pount you have Mac application that you can run from any location on Mac
You can add this app into crontab by executing
   crontab -e 
and adding following line
* * * * * /path_to_PurgeSpam_app/PurgeSpam.app/Contents/MacOS/applet
