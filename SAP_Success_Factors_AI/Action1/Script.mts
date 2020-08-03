'===========================================================
'Make sure you have your VM set to full screen, not windowed mode
'If the application is no longer available, contact the SAP SWAT team member appropriate for your region to get a new one provisioned:
'	North America - Don Jackson
'	International - Jan DeCoster
'	APJ - Bo Hee Seo
'===========================================================


'===========================================================
'Function to Create a Random Number with DateTime Stamp
'===========================================================
Function fnRandomNumberWithDateTimeStamp()

'Find out the current date and time
Dim sDate : sDate = Day(Now)
Dim sMonth : sMonth = Month(Now)
Dim sYear : sYear = Year(Now)
Dim sHour : sHour = Hour(Now)
Dim sMinute : sMinute = Minute(Now)
Dim sSecond : sSecond = Second(Now)

'Create Random Number
fnRandomNumberWithDateTimeStamp = Int(sDate & sMonth & sYear & sHour & sMinute & sSecond)

'======================== End Function =====================
End Function

'===========================================================
'Function for debugging properties at run time to output to the log
'===========================================================
Function PropertiesDebug
'Debug code to determine why the checkpoint was failing, turns out that there is a trailing space in the application code that the result HTML was trimming when displaying expected vs. actual
'CPExpected = "'" & DataTable.Value ("FullName") & "'"										'Set the variable for what is in the data table, enclose with single quotes so we can find leading/trailing spaces
'CPActual = Browser("Browser").Page("SuccessFactors: Candidates").Link("CandidateName").GetROProperty("text")	'Get the actual text from the object at run time
'CPActual = "'" & CPActual & "'"															'Set the variable for what is the object property at run time enclosed with a single quotes so we can find leading/trailing spaces
'Print "Expected is " & CPExpected															'Output the expected value to the output log
'Print "Actual is " & CPActual																'Output the actual value to the output log
	
End Function

Dim FirstName, LastName, Email, CPActual, CPExpected

Browser("Browser").Maximize																	'Maximize the browser window
Browser("Browser").Navigate DataTable.GlobalSheet.GetParameter("URL")						'Navigate to the URL of SuccessFactors, driven off of the datasheet
AIUtil.SetContext Browser("Browser")														'Tell the AI SDK which window to work against
'===========================================================================================
'The grey on blue text is not being recognized for the Username text consistently due to the lack of contrast.
'	As such, you can use the icon next to the Username field to find the Username input field
'===========================================================================================
Set IconAnchor = AIUtil("profile")															'Set the IconAnchor to be the profile icon
Set ValueAnchor = AIUtil("input", micAnyText, micWithAnchorOnLeft, IconAnchor)				'Set the Value field to be an "input" field, with any text, with the IconAnchor to its left
ValueAnchor.Type DataTable.GlobalSheet.GetParameter("Username")								'Enter the User Name from the datasheet into the username field
AIUtil("text_box", "Enter Password").Type DataTable.GlobalSheet.GetParameter("Password")	'Enter the password fromt he datasheet into the Password field
AIUtil.FindTextBlock("Log in").Click														'Click the Login button
Browser("Browser").Sync																		'Wait for the browser DOM to be ready to proceed
'===========================================================================================
'You can just use AIUtil("down_triangle", micNoText, micFromLeft, 1).Click to click the down arrow,
'	but it is easier to understand/read if you anchor off of the text "Admin Center".  Please note that
'	there are two text blocks with that text, so it's the first from the top
'===========================================================================================
Set TextAnchor = AIUtil.FindTextBlock("Admin Center", micFromTop, 1)
Set IconAnchor = AIUtil("down_triangle", micNoText, micWithAnchorOnLeft, TextAnchor)
IconAnchor.Click
'===========================================================================================
'AI ADK OCR isn't recognizing the menu, submitted using feedback tool
'===========================================================================================
Browser("Browser").Page("SuccessFactors: Admin").Link("Recruiting").WaitProperty "visible",True, 3000	'Wait for the application to be ready to proceed
Browser("Browser").Page("SuccessFactors: Admin").Link("Recruiting").Click					'The AI SDK OCR can't see the Recruiting text on the screen, occasionally
AIUtil.FindTextBlock("Preferences").Exist 													'Sync to ensure the menu loads before proceeding.
AIUtil.FindTextBlock("Candidates", micFromTop, 1).Click										'Click the Candidates tab item at the top of the screen
AIUtil.FindText("Add Candidate").Click														'Click the Add Candidate text
'===========================================================================================
'The application sometimes will shift the location of the Add Candidate text as it is loading,
'	in between when the object is found on the screen, and the click occurs, as such 
'	check to make sure the subsequent statement comes up, if not, re-do the Add Candidate click
'===========================================================================================
If AIUtil("text_box", "First Name").Exist = False Then
	AIUtil.FindText("Add Candidate").Click														'Click the Add Candidate text
End If
FirstName = "FN" & fnRandomNumberWithDateTimeStamp											'Create a unique first name
AIUtil("text_box", "First Name").Type FirstName												'Enter the first name into the First Name field
LastName = "LN" & fnRandomNumberWithDateTimeStamp											'Create a unique last name
AIUtil("text_box", "Last Name").Type LastName												'Enter the last name into the Last Name field
Email = "email" & fnRandomNumberWithDateTimeStamp & DataTable.GlobalSheet.GetParameter("EmailDomain")	'Create a unique e-mail address, using the e-mail domain in the data sheet
AIUtil("text_box", "Email:").Type Email														'Enter the e-mail into the e-mail address field
Browser("Browser").Page("SuccessFactors: Candidates").WebElement("Country/Region:").Click	'AI SDK Currently doesn't have a scroll command, so leverage traditional OR to force the browser to scroll
AIUtil("text_box", "Retype Email Address").Type Email										'Enter the e-mail into the re-enter e-mail address field
AIUtil("text_box", "Phone").Type DataTable.GlobalSheet.GetParameter("PhoneNumber")			'Enter in a phone number into the phone number field from the data table
Country = DataTable.GlobalSheet.GetParameter("Country")										'Set the variable to be the value for Country from the data sheet
Browser("Browser").Page("SuccessFactors: Candidates").WebList("select").Select Country		'Select the Country with the value from the data sheet, AI SDK ComboBox doesn't work on this particular combo box, use traditional OR @@ script infofile_;_ZIP::ssf4.xml_;_
AIUtil("button", "Create Profile").Click													'Click the Create Profile button
AIUtil.SetContext Browser("Candidate Profile")												'Tell the AI SDK to start working against the pop-up browser window
Browser("Candidate Profile").Maximize														'Maximize the pop-up browser window
AIUtil("button", "Cancel").Click															'Click the Cancel button to not upload a resume
AIUtil.FindTextBlock("+ Add", micFromTop, 1).Click											'Click the Add button to add internal job history
AIUtil("button", "Close Details").Click														'Change your mind, don't enter internal job history and slide down table
AIUtil("combobox", "Salutation").Select "Mr."												'Select the Mr. salutation
AIUtil.FindTextBlock("Save").Click															'Save changes in the candidate pop-up browser window
AIUtil.FindTextBlock("Close Window").Click													'Close the pop-up browser window with the Close Window text link
AIUtil.SetContext Browser("Browser")														'Tell the AI SDK to work against the initial window again
AIUtil("close", micNoText, micFromRight, 1).Click											'Remove the other (default) search criteria line
AIUtil.FindText("Basic Info").Click															'Click the Basic Info drop down to be able to search by the name
AIUtil.FindTextBlock("First Name").Click													'Click the First Name in the drop down
AIUtil.FindText("Marketing Emails").Click													'Click this line to shift focus, and erase the visual element of the box around the new line
'===========================================================================================
'We want to find the line that has "First Name" to find the text box to type in the first name
'	This is another example of using VRI to find the correct object.
'===========================================================================================
AIUtil.SetContext Browser("Browser")														'Tell the AI SDK to work against the initial window again
Set LineAnchor = AIUtil.FindText("First Name")												'Find the text "First Name" on the screen, set that as the line anchor
Set ValueAnchor = AIUtil("text_box", micAnyText, micWithAnchorOnLeft, LineAnchor)			'Set the ValueAnchor to be the text box with the line anchor on the left
ValueAnchor.Type FirstName																	'Enter the same first name for the candidate created earlier
AIUtil("button", "", micFromBottom, 1).Click												'Click the Search button
AIUtil("button", "Accept").Click															'Click the Accept button on the pop-up frame to accept search results
DataTable.Value ("FullName") = FirstName & " " & LastName & " "								'Set the value in the data table for the calculated full name of the candidate, used in the next step
Browser("Browser").Page("SuccessFactors: Candidates").Link("CandidateName").Check CheckPoint("CPCandidateFullName")	'Checkpoint to make sure that the candidate link showed up @@ script infofile_;_ZIP::ssf18.xml_;_
Browser("Browser").Page("SuccessFactors: Candidates").SAPUIButton("Account Navigation for").Click	'There isn't anything for AI to recognize for the user drop down, it's a picture of the person, use traditional OR @@ script infofile_;_ZIP::ssf5.xml_;_
AIUtil.FindText("Log out").Click															'Click the Log out text in the drop down menu
AIUtil.FindTextBlock("Log in").Exist														'Wait for the Login text to make sure the window has finished loading
Browser("Browser").Close																	'Close the browser

