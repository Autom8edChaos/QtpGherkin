' Demonstration of a gherkin language interpreter 
Option Explicit

' Register givens with the given attribute. Multiple givens (or even whens or thens) can
' be registered to a single function.
' The parameter is wildcarded as a regular expression and will be matched
' The exact functionname to register to is still necessary unfortunately			
[given] "a client '(.*)'", "SelectClient"
public function SelectClient(params)
	msgbox "Selecting a client with id " & params(0)
end function

' The when attribute, works the same as the given attribute
[when] "I add an account", "AddAnAccount"
public function AddAnAccount(params)
	msgbox "Adding an account"
end function

' Demonstration of a then attribute
[then] "I should receive a new account", "ReceiveNewAccount"
public function ReceiveNewAccount(params)
	msgbox "Receiving a new account"
end function

' Then attribute with multiple parameters
[then] "a new account number with format '(.*)' with saldo '(.*)'", "CheckAccountFormatAndSaldo"
public function CheckAccountFormatAndSaldo(params)
	msgbox "Check an account with number format " & params(0)
	msgbox "Check a saldo on this account of " & params(1)
end function

' The featuretext can be loaded by a text file, for demo purposes in an array
dim featureText
featureText = array("Scenario: A new account can be added to a client", _
			"Given a client 'testklant'", _
			"When I add an account", _
			"Then I should receive a new account", _
			"And a new account number with format '\d+' with saldo '0'", _
			"", _
			"Scenario: Another account can be added to the same client", _
			"Given a client 'testklant'", _
			"When I add an account", _
			"Then I should receive a new account", _
			"And a new account number with format '\d+' with saldo '0'")
			
runFeature featureText

' Feature runner. The call to this runner function must occur _after_ the registration
' of the given/when/thens. For easy demo purposes this function accepts an array.
private sub runFeature(fArray)			
	dim command
	with new cls_GherkinCommandParser
		.GherkinTextArray = fArray
		for each command in .GetCommandStream
			command.Exec()
		next
	end with
end sub 




' Implementation down here
' ------------------------
Class cls_GherkinCommandParser

	Public GherkinTextArray
	Public CommandDict
	
	Public Function GetCommandStream()
		dim stream, line, command
		Set stream = CreateObject("System.Collections.ArrayList")
		
		For each line in GherkinTextArray
			Set command = PhraseCollection.FindMatchingCommand(line)
			If Not (command Is Nothing) Then
				stream.Add command
			End if
		Next
		GetCommandStream = stream.ToArray()
	End Function
	
End Class

public sub [given](phrase, functionString)
	PhraseCollection.Add "Given", phrase, functionString
end sub
public sub [when](phrase, functionString)
	PhraseCollection.Add "When", phrase, functionString
end sub
public sub [then](phrase, functionString)
	PhraseCollection.Add "Then", phrase, functionString
end sub

Private STATIC_PhraseCollection 
Public Function PhraseCollection()
	if not isObject(STATIC_PhraseCollection) then
		Set STATIC_PhraseCollection = new cls_PhraseCollection
	end if
	Set PhraseCollection = STATIC_PhraseCollection
End Function

Class cls_PhraseCollection

	Private Sub Class_Initialize()
		Set commandObjectCollection = CreateObject("Scripting.Dictionary")
	End Sub
	Public commandObjectCollection
	Public Sub Add(gType, phrase, functionString)
		commandObjectCollection.Add gType & "::" & phrase, new_CommandObject(lCase(gType), phrase, functionString)
	End Sub
	
	private lastAction_
	Public function FindMatchingCommand(line)
		dim words, gtype, commandObject, matches, keyword
		dim smCount, params, i
		Set FindMatchingCommand = Nothing
		dim re : set re = new RegExp
		re.Global = True
		re.IgnoreCase = True
				
		words = split(line, " ")
		if ubound(words) <= 0 then Exit Function
		
		keyword = lcase(words(0))
		
		' See if the keyword is a special
		Select case keyword
			case "scenario:"
				Set FindMatchingCommand = new_scenarioCommand(line)
			
			case else
				for each commandObject in commandObjectCollection.Items
					if keyword = "and" then keyword = lastAction_
					
					if commandObject.gType = keyword then
						lastAction_ = keyword
						
						re.Pattern = commandObject.Phrase & "$"
						If re.Test(line) then
							Set matches = re.Execute(line)
							smCount = matches(0).SubMatches.count
							if smCount > 0 then
								redim params(smCount-1)
								for i = 0 to smCount-1
									params(i) = matches(0).SubMatches(i)
								Next
							else
								params = array()
							End if
							commandObject.params = params
							Set FindMatchingCommand = commandObject
							Exit for
						End if
					End if			
				next
		End select
			
	end function

End Class

Public Function new_CommandObject(gType, phrase, functionString)
	Dim tObj : Set tObj = new cls_CommandObject
	tObj.gType = gType
	tObj.Phrase = phrase
	tObj.FunctionString = functionString
	Set new_CommandObject = tObj
End Function

Class cls_CommandObject
	Public gType, phrase
	private functionString_, functionPointer_
	Public Property Let FunctionString(fs)
		functionString_ = fs
		Set functionPointer_ = GetRef(fs)
	End Property
	
	Public Property Get FunctionPointer()
		Set FunctionPointer = functionPointer_
	End Property
	Public Property Get FunctionString()
		FunctionString = functionString_
	End Property
		
	Public params
	
	Public Function Exec()
		Call FunctionPointer()(params)
	End Function
End Class

Public Function new_ScenarioCommand(text)
	Dim tObj : Set tObj = new cls_GenericReporter
	tObj.text = text
	Set new_ScenarioCommand = tObj
End Function

Class cls_GenericReporter

	Public text
	
	Public Function Exec()
		Msgbox text
	End Function
	
End Class

