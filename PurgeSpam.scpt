#AppleScript to delete emails of specific sender from Trash folder of Microsoft Outllok on Mac
# Some initial variables
set today to (current date)
set cutoff to 2 * days
set shortCutoff to 3 * hours

tell application "Microsoft Outlook"

  log "Started"
  set myTrash to folder "Deleted items" of default account
  
  set theMessages to messages of myTrash

  # A couple of counters used for logging changes
  set countMoved to 0
  set countDeleted to 0

  repeat with theMessage in theMessages
    try
      
      set theSender to sender of theMessage
      set fromAddress to address of theSender      
      # Start of mail specific rules
      
      # This deletes a CI notification if it is more than 3 hours old or has been read
      if fromAddress is "sender1@domain.com" or fromAddress is "sender2@domain.com" then        
          delete theMessage
          set countDeleted to countDeleted + 1
      end if
      

      # and so on for every rule

      # End of mail specific rules

    on error errorMsg
      log "Error: " & errorMsg
    end try
  end repeat
  
  log "Outlook cleanup ran at " & today & " - " & countMoved & " moved, " & countDeleted & " deleted"
end tell
