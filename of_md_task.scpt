(*
File: OmniFocus_Due_List.scpt
-------------------------------------------------------------------------------------------------
Revision: 4.15
Revised: 2019-04-04
Summary: Create .md file for list of tasks due and deferred +/- 7d from current date.
-------------------------------------------------------------------------------------------------
Script based on Justin Lancy (@veritrope) from Veritrope.com
http://veritrope.com/code/write-todays-completed-tasks-in-omnifocus-to-a-text-file
-------------------------------------------------------------------------------------------------
*)

--Set Date Functions
set CurrDatetxt to short date string of date (short date string of (current date))
set dateYeartxt to year of (current date) as integer

if (month of (current date) as integer) < 10 then
	set dateMonthtxt to "0" & (month of (current date) as integer)
else
	set dateMonthtxt to month of (current date) as integer
end if

if (day of (current date) as integer) < 10 then
	set dateDaytxt to "0" & (day of (current date) as integer)
else
	set dateDaytxt to day of (current date) as integer
end if

set str_date to "" & dateYeartxt & "-" & dateMonthtxt & "-" & dateDaytxt

--Set File/Path name of MD file
--set theFilePath to choose file name default name "To Do List for " & str_date & ".md"
set theFilePath to ((path to desktop folder) as string) & "To Do List for " & str_date & ".md"

--Get OmniFocus task list
set due_Tasks to my OmniFocus_task_list()

--Output .MD text file
my write_File(theFilePath, due_Tasks)

--Set OmniFocus Due Task List
on OmniFocus_task_list()
	set endDate to (current date) + (7 * days)
	set startDate to (current date) - (14 * days)
	set CurrDate to date (short date string of (startDate))
	set CurrDatetxt to short date string of date (short date string of (current date))
	set endDatetxt to date (short date string of (endDate))
	tell application "OmniFocus"
		tell default document
			set refDueTaskList to a reference to (flattened tasks where (due date < endDatetxt and completed = false))
			set {lstName, lstProject, lstContext, lstDueDate} to {name, name of its containing project, name of its primary tag, due date} of refDueTaskList
			set strText to "To Do List for " & CurrDatetxt & ":" & return & return
			repeat with iTask from 1 to count of lstName
				set {strName, varProject, varContext, varDueDate} to {item iTask of lstName, item iTask of lstProject, item iTask of lstContext, item iTask of lstDueDate}
				if (varDueDate < (current date)) then
					set strDueDate to "<span style=\"color:red\">" & short date string of varDueDate & "</span>"
				else
					set strDueDate to short date string of varDueDate
				end if
				set strText to strText & "▢ " & strName & " " & strDueDate
				set strText to strText & return
			end repeat
		end tell
		
		tell default document
			set ref2DueTaskList to a reference to (flattened tasks where (defer date < endDatetxt and (due date < CurrDate or due date is missing value) and completed = false))
			set {lst2Name, lst2Project, lst2Context, lst2DeferDate, lst2DueDate} to {name, name of its containing project, name of its primary tag, defer date, due date} of ref2DueTaskList
			set str2Text to return & return & "Tasks to start this week:" & return & return
			repeat with i2Task from 1 to count of lst2Name
				set {str2Name, var2Project, var2Context, var2DeferDate, var2DueDate} to {item i2Task of lst2Name, item i2Task of lst2Project, item i2Task of lst2Context, item i2Task of lst2DeferDate, item i2Task of lst2DueDate}
				if (var2DeferDate < (current date)) then
					set str2DueDate to "<span style=\"color:blue\">" & short date string of var2DeferDate & "</span>"
				else
					set str2DueDate to short date string of var2DeferDate
				end if
				set str2Text to str2Text & "∙ " & str2Name & " " & str2DueDate
				set str2Text to str2Text & return
			end repeat
		end tell
	end tell
	set str3Text to strText & str2Text
	str3Text
end OmniFocus_task_list

--Export Task list to .MD file
on write_File(theFilePath, due_Tasks)
	set theText to due_Tasks
	set theFileReference to open for access theFilePath with write permission
	write theText to theFileReference as «class utf8»
	close access the theFileReference
end write_File


tell application "Marked"
	open file theFilePath
end tell
