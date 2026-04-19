#!/usr/bin/osascript
-- ─────────────────────────────────────────────────────────────────────────────
-- Keep Hammering — Apple Notes Weekly Generator
-- Creates 7 daily Apple Notes (Mon–Sun) in STM > Daily Notes folder.
--
-- Run manually:   osascript ~/Documents/KeepHammering/create_apple_notes.applescript
-- Auto-run:       Triggered every Monday at 6:00 AM via Shortcuts automation
-- ─────────────────────────────────────────────────────────────────────────────


-- ── Date Helpers ──────────────────────────────────────────────────────────────

-- Returns the Monday of the week containing the given date
on getMondayOf(d)
	set dayInt to weekday of d as integer
	-- AppleScript weekday integers: 1=Sun 2=Mon 3=Tue 4=Wed 5=Thu 6=Fri 7=Sat
	if dayInt = 1 then
		set dayOffset to -6 -- Sunday → back 6 days to previous Monday
	else
		set dayOffset to -(dayInt - 2)
	end if
	return d + (dayOffset * days)
end getMondayOf

-- Returns "Monday, 23 March 2026"
on formatDateLong(d)
	set dayNames to {"Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"}
	set monthNames to {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
	set dName to item (weekday of d as integer) of dayNames
	set mName to item (month of d as integer) of monthNames
	return dName & ", " & (day of d) & " " & mName & " " & (year of d)
end formatDateLong

-- Returns "#2026-03" style month tag
on getMonthTag(d)
	set y to (year of d) as string
	set mInt to month of d as integer
	if mInt < 10 then
		set mm to "0" & mInt
	else
		set mm to mInt as string
	end if
	return "#" & y & "-" & mm
end getMonthTag


-- ── HTML / TODO Helpers ──────────────────────────────────────────────────────

-- Strip all HTML tags from a string, returning plain text
on stripHTML(htmlStr)
	set cleanStr to ""
	set inTag to false
	repeat with c in (characters of htmlStr)
		set ch to c as string
		if ch is "<" then
			set inTag to true
		else if ch is ">" then
			set inTag to false
		else if not inTag then
			set cleanStr to cleanStr & ch
		end if
	end repeat
	return cleanStr
end stripHTML

-- Extract unchecked, non-empty TODO checklist items from a note body (HTML string)
on extractTodoItems(noteBody)
	set todoItems to {}
	set oldDelims to AppleScript's text item delimiters

	-- Isolate the first <ul class="checklist">…</ul> block
	set AppleScript's text item delimiters to "<ul class=\"checklist\">"
	set parts to text items of noteBody
	if (count of parts) < 2 then
		set AppleScript's text item delimiters to oldDelims
		return {}
	end if
	set afterStart to item 2 of parts

	set AppleScript's text item delimiters to "</ul>"
	set checklistContent to item 1 of (text items of afterStart)

	-- Walk each <li … > chunk
	set AppleScript's text item delimiters to "<li"
	set liParts to text items of checklistContent

	repeat with i from 2 to count of liParts
		set liChunk to item i of liParts

		-- Split on first ">" to separate tag attributes from inner content
		set AppleScript's text item delimiters to ">"
		set tagParts to text items of liChunk
		set openTag to item 1 of tagParts -- attributes, empty if bare <li>

		-- Skip checked items (Apple Notes marks them with "checked" in the tag)
		if openTag does not contain "checked" then
			if (count of tagParts) >= 2 then
				-- Reassemble everything after the first ">"
				set liInner to ""
				repeat with j from 2 to count of tagParts
					if j > 2 then set liInner to liInner & ">"
					set liInner to liInner & (item j of tagParts as string)
				end repeat

				-- Grab content before </li>
				set AppleScript's text item delimiters to "</li>"
				set liContent to item 1 of (text items of liInner)

				set itemText to my stripHTML(liContent)
				if itemText is not "" then
					set todoItems to todoItems & {itemText}
				end if
			end if
		end if
	end repeat

	set AppleScript's text item delimiters to oldDelims
	return todoItems
end extractTodoItems


-- ── Note Body Builder ─────────────────────────────────────────────────────────

-- carryItems: list of strings to pre-populate in Today's TODO (Monday only)
on buildDayNote(dateLabel, mTag, carryItems)

	-- ── Title / Date (bold — this becomes the note title automatically) ────────
	set n to "<div><b><i><u><span style=\"font-size:16pt;\">" & dateLabel & "</span></u></i></b></div>"
	set n to n & "<div><br></div>"

	-- ── Motto ─────────────────────────────────────────────────────────────────
	set n to n & "<div style=\"text-align:center;\"><b><i><span style=\"color:#CC0000;\">Aut Viam Inveniam Aut Faciam</span></i></b></div>"
	set n to n & "<div><br></div>"

	-- ── Customers Seen ────────────────────────────────────────────────────────
	set n to n & "<div><b><i><span style=\"color:#CC0000;\">– Focus on income // Self improvement*</span></i></b></div>"
	set n to n & "<div><b><u>Customers Seen</u></b></div>"
	set n to n & "<ul>"
	set n to n & "<li></li>"
	set n to n & "<li></li>"
	set n to n & "<li></li>"
	set n to n & "</ul>"
	set n to n & "<div><br></div>"

	-- ── Today's TODO (native checkboxes) ──────────────────────────────────────
	set n to n & "<div><b><i><u>Today's TODO</u></i></b></div>"
	if (count of carryItems) > 0 then
		set n to n & "<div><i><span style=\"color:#999999;\">&#8629; Carried from Friday:</span></i></div>"
	end if
	set n to n & "<ul class=\"checklist\">"
	if (count of carryItems) > 0 then
		repeat with carryItem in carryItems
			set n to n & "<li>" & carryItem & "</li>"
		end repeat
		-- Keep three empty slots for new items
		set n to n & "<li></li>"
		set n to n & "<li></li>"
		set n to n & "<li></li>"
	else
		set n to n & "<li></li>"
		set n to n & "<li></li>"
		set n to n & "<li></li>"
		set n to n & "<li></li>"
		set n to n & "<li></li>"
	end if
	set n to n & "</ul>"
	set n to n & "<div><br></div>"

	-- ── Personal TODO (native checkboxes) ────────────────────────────────────
	set n to n & "<div><b><i><u>Personal TODO</u></i></b></div>"
	set n to n & "<ul class=\"checklist\">"
	set n to n & "<li></li>"
	set n to n & "<li></li>"
	set n to n & "<li></li>"
	set n to n & "<li></li>"
	set n to n & "<li></li>"
	set n to n & "</ul>"
	set n to n & "<div><br></div>"

	-- ── Notes ─────────────────────────────────────────────────────────────────
	set n to n & "<div><b><i><u>Notes:</u></i></b></div>"
	set n to n & "<ol>"
	set n to n & "<li></li>"
	set n to n & "<li></li>"
	set n to n & "<li></li>"
	set n to n & "<li></li>"
	set n to n & "<li></li>"
	set n to n & "</ol>"
	set n to n & "<div><br></div>"

	-- ── Scratch Pad ───────────────────────────────────────────────────────────
	set n to n & "<div><b><i><u>Scratch pad:</u></i></b></div>"
	set n to n & "<div><i><u>The best way out is always through.</u></i></div>"
	set n to n & "<div style=\"padding-left:40px;\"><i><u>-Robert Frost</u></i></div>"
	set n to n & "<div><br></div>"
	set n to n & "<div><br></div>"
	set n to n & "<div><br></div>"
	set n to n & "<div><br></div>"
	set n to n & "<div><br></div>"
	set n to n & "<div><br></div>"
	set n to n & "<div><br></div>"

	-- ── Bottom Banner ─────────────────────────────────────────────────────────
	set n to n & "<div style=\"text-align:center;\"><b><i><span style=\"color:#CC0000; font-size:16pt;\">Keep Hammering</span></i></b></div>"
	set n to n & "<div><br></div>"

	-- ── Tags ──────────────────────────────────────────────────────────────────
	set n to n & "<div><i><span style=\"color:#999999;\">#daily  #work  " & mTag & "  (add: #meetings #ideas #urgent #wins as needed)</span></i></div>"

	return n
end buildDayNote


-- ── Main ──────────────────────────────────────────────────────────────────────

set today to current date
set weekStart to my getMondayOf(today)

-- Last Friday = Monday minus 3 days
set lastFriday to weekStart - (3 * days)
set fridayLabel to my formatDateLong(lastFriday)
set carryForwardItems to {}

tell application "Notes"
	set targetAccount to default account

	-- ── Find STM > Daily Notes folder ─────────────────────────────────────────
	set stmFolder to missing value
	set targetFolder to missing value

	tell targetAccount
		-- Find the top-level STM folder
		repeat with f in folders
			if name of f is "STM" then
				set stmFolder to f
				exit repeat
			end if
		end repeat

		if stmFolder is missing value then
			error "Could not find a folder named 'STM' in Notes. Check the folder name and try again."
		end if

		-- Find Daily Notes inside STM
		repeat with sf in folders of stmFolder
			if name of sf is "Daily Notes" then
				set targetFolder to sf
				exit repeat
			end if
		end repeat

		if targetFolder is missing value then
			error "Could not find 'Daily Notes' inside STM. Check the subfolder name and try again."
		end if

		-- ── Read last Friday's TODO items ─────────────────────────────────────
		repeat with existingNote in notes of targetFolder
			if name of existingNote is fridayLabel then
				set carryForwardItems to my extractTodoItems(body of existingNote)
				exit repeat
			end if
		end repeat

		-- ── Create one note per day (Monday through Sunday) ───────────────────
		repeat with i from 0 to 6
			set noteDate to weekStart + (i * days)
			set dateLabel to my formatDateLong(noteDate)
			set mTag to my getMonthTag(noteDate)

			-- Skip if a note with this title already exists (prevents duplicates)
			set alreadyExists to false
			repeat with existingNote in notes of targetFolder
				if name of existingNote is dateLabel then
					set alreadyExists to true
					exit repeat
				end if
			end repeat

			if not alreadyExists then
				-- Pass carry-forward items only to Monday (i = 0)
				if i = 0 then
					set noteBody to my buildDayNote(dateLabel, mTag, carryForwardItems)
				else
					set noteBody to my buildDayNote(dateLabel, mTag, {})
				end if
				-- Body only — no name property. Apple Notes derives the title
				-- from the first line of the body (the bold date heading).
				make new note at targetFolder with properties {body: noteBody}
			end if
		end repeat

	end tell
end tell

-- ── Open Notes so banner images can be added in one focused pass ──────────────
-- All 7 notes are created. Notes will open — drag the banner image into each
-- note now while you're here. Takes about 1 minute to do all 7 at once.
tell application "Notes"
	activate
end tell
