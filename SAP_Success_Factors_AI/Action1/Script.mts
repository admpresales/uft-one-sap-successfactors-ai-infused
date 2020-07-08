'===========================================================
'Make sure you have your VM set to full screen, not windowed mode
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
'The grey on blue text is not being recognized for the Username text consistently, use traditional OR
Browser("Browser").Page("SuccessFactors Log in").WebEdit("username").Set DataTable.GlobalSheet.GetParameter("Username")				'Enter the User Name from the datasheet into the username field
AIUtil("text_box", "Enter Password").Type DataTable.GlobalSheet.GetParameter("Password")	'Enter the password fromt he datasheet into the Password field
AIUtil.FindTextBlock("Log in").Click														'Click the Login button
Browser("Browser").Sync																		'Wait for the browser DOM to be ready to proceed
AIUtil("down_triangle", micNoText, micFromLeft, 1).Highlight								'Highlight the down triangle for the menu
AIUtil("down_triangle", micNoText, micFromLeft, 1).Click									'Click the down triangle for the menu
Browser("Browser").Page("SuccessFactors: Admin").Link("Recruiting").WaitProperty "visible",True, 3000	'Wait for the application to be ready to proceed
Browser("Browser").Page("SuccessFactors: Admin").Link("Recruiting").Click					'The AI SDK is struggling with the Recruiting text on the screen, occasionally at run time, it can't find it
Browser("Browser").Page("SuccessFactors: Job Requisitio").WebTable("Job Requisition Summary").WaitProperty "visible",True, 3000	'Wait for the application to be ready to proceed, keying off of the Job Requisition data table, using traditional OR
AIUtil.FindTextBlock("Candidates", micFromTop, 1).Click										'Click the Candidates tab item at the top of the screen
Browser("Browser").Page("SuccessFactors: Candidates").WebElement("AddCandidatePlusButton").Click	'Click the Add Candidate button, AIUtil currently can't see the + button well
FirstName = "FN" & fnRandomNumberWithDateTimeStamp											'Create a unique first name
AIUtil("text_box", "First Name").Type FirstName												'Enter the first name into the First Name field
LastName = "LN" & fnRandomNumberWithDateTimeStamp											'Create a unique last name
AIUtil("text_box", "Last Name").Type LastName												'Enter the last name into the Last Name field
Email = "email" & fnRandomNumberWithDateTimeStamp & DataTable.GlobalSheet.GetParameter("EmailDomain")	'Create a unique e-mail address, using the e-mail domain in the data sheet
AIUtil("text_box", "Email:").Type Email														'Enter the e-mail into the e-mail address field
Browser("Browser").Page("SuccessFactors: Candidates").WebElement("Country/Region:").Click	'AI SDK Currently doesn't have a scroll command, so leverage traditional OR to force the browser to scroll
AIUtil("text_box", "Retype Email Address").Type Email												'Enter the e-mail into the re-enter e-mail address field
AIUtil("text_box", "Phone").Type DataTable.GlobalSheet.GetParameter("PhoneNumber")			'Enter in a phone number into the phone number field from the data table
Country = DataTable.GlobalSheet.GetParameter("Country")										'Set the variable to be the value for Country from the data sheet
Browser("Browser").Page("SuccessFactors: Candidates").WebList("select").Select Country		'Select the Country with the value from the data sheet, new AI SDK ComboBox doesn't work on this particular combo box, use traditional OR @@ script infofile_;_ZIP::ssf4.xml_;_
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
AIUtil.FindTextBlock("Basic Info V").Click													'Click the Basic Info drop down to be able to search by the name
AIUtil.FindTextBlock("First Name").Click													'Click the First Name in the drop down

Browser("Browser").Page("SuccessFactors: Candidates").WebEdit("FirstNameSearchBox").Highlight	'Enter the same first name for the candidate created earlier
Browser("Browser").Page("SuccessFactors: Candidates").WebEdit("FirstNameSearchBox").Set FirstName	'Enter the same first name for the candidate created earlier

'====================================================================================================
'	This is a section to work on building VRI to click in the search box rather than using traditional OR

'AIUtil.SetContext Browser("Browser")														'Tell the AI SDK which window to work against
'
''Set AnchorObject = AIUtil.FindText("Profile", micFromTop, 2) 'Describe the anchor object
'Set AnchorObject = AIUtil.FindText("and First Name")
'AnchorObject.Highlight
''AIUtil.FindText("User Name",micWithAnchorOnLeft, AnchorObject).Type "FN145202084557"
'AIUtil.(("text_box"),micWithAnchorOnLeft, AnchorObject).Type "FN145202084557"
'AIUtil.FindTextBlock("and First Name").Click
'
'
'AIUtil("combobox", "and First Name").Highlight


''The following example clicks on the user name to the right of the 2nd "Profile"
'' text from the top. Then it clicks on the login text.
'AIUtil.SetContext Device("device")  'Set the context for AI
'Set secondProfileFromTop = AIUtil.FindText("Profile", micFromTop, 2) 'Describe the anchor object
'AIUtil.FindText("User Name",micWithAnchorOnLeft, secondProfileFromTop).Click
'AIUtil.FindTextBlock("Log In With Your App").Click 10, 10 'Click the 10,10 point in the LogIn text

'AIUtil("text_box", "", micFromBottom, 1).Highlight											'There is a need to sometimes slow the script to allow the app to catch up
'AIUtil("text_box", "", micFromBottom, 1).Type FirstName										'Enter the same first name for the candidate created earlier
'====================================================================================================

Browser("Browser").Page("SuccessFactors: Candidates").WebButton("Search").Click				'Click the Search button
AIUtil("button", "Accept").Click															'Click the Accept button on the pop-up frame to accept search results
DataTable.Value ("FullName") = FirstName & " " & LastName & " "								'Set the value in the data table for the calculated full name of the candidate, used in the next step
Browser("Browser").Page("SuccessFactors: Candidates").Link("CandidateName").Check CheckPoint("CPCandidateFullName")	'Checkpoint to make sure that the candidate link showed up @@ script infofile_;_ZIP::ssf18.xml_;_
Browser("Browser").Page("SuccessFactors: Candidates").SAPUIButton("Account Navigation for").Click	'There isn't anything for AI to recognize for the user drop down, it's a picture of the person, use traditional OR @@ script infofile_;_ZIP::ssf5.xml_;_
AIUtil.FindTextBlock("Q) Log out").Click													'Click the Log out text in the drop down menu
AIUtil.FindTextBlock("Log in").Highlight													'Highlight the Login text to make sure the window has finished loading
Browser("Browser").Close																	'Close the browser

