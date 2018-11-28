#AppleScript to delete emails of specific sender from Trash folder of Microsoft Outllok on Mac

tell application "Microsoft Outlook"

  log "Started"
  set myTrash to folder "Deleted items" of default account
  
  set theMessages to messages of myTrash

  repeat with theMessage in theMessages
    try
      
      set theSender to sender of theMessage
      set fromAddress to address of theSender   
      
      if fromAddress is "sender1@domain.com" or fromAddress is "sender2@domain.com" then        
          delete theMessage
      end if     

    on error errorMsg
      log "Error: " & errorMsg
    end try
  end repeat
end tell
