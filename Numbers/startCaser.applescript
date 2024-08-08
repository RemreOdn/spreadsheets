tell application "Numbers"
	tell front document
		tell active sheet
			tell (first table whose selection range's class is range)
				set selectedRange to selection range
				repeat with aCell in selectedRange's cells
					set currentValue to value of aCell
					if currentValue is not missing value then
						set newValue to do shell script "echo " & quoted form of currentValue & " | awk '{for(i=1;i<=NF;i++) $i=toupper(substr($i,1,1)) tolower(substr($i,2))}1'"
						set value of aCell to newValue
					end if
				end repeat
			end tell
		end tell
	end tell
end tell
