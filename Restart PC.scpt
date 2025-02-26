try
	set dialogResult to display dialog "Please save your work. The computer will restart." buttons {"Cancel", "Restart"} default button "Restart" with icon caution giving up after 120
	
	if the gave up of the dialogResult is true then
		display dialog "Restart canceled due to timeout." buttons {"OK"} default button "OK" with icon note
	else if the button returned of the dialogResult is "Cancel" then
		display dialog "Restart canceled by user." buttons {"OK"} default button "OK" with icon note
	else if the button returned of the dialogResult is "Restart" then
		
		-- Prevent macOS from reopening apps after restart
		do shell script "defaults write com.apple.loginwindow TALLogoutSavesState -bool false; killall cfprefsd"
		
		-- Restart the computer
		do shell script "osascript -e 'tell application \"System Events\" to restart'"
	end if
on error errMsg number errNum
	if errNum = -128 then
		display dialog "Restart canceled by the user." buttons {"OK"} default button "OK" with icon note
	end if
end try
return

