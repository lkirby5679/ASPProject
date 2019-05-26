<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  forum code based on Ti Portal's forum sample.  http://www.transworldportal.com '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const forums_forum_id = 0
Const forums_forum_name = 1
Const forums_forum_description = 2
Const forums_forum_start_date = 3
Const forums_forum_grouping = 4

Const messages_message_id = 0
Const messages_forum_id = 1
Const messages_thread_id = 2
Const messages_thread_parent = 3
Const messages_thread_level = 4
Const messages_message_author = 5
Const messages_message_author_email = 6
Const messages_message_author_notify = 7
Const messages_message_timestamp = 8
Const messages_message_subject = 9
Const messages_message_body = 10
Const messages_message_Approved = 11

Function GetMessages(iActiveMessageID)
	Dim objRSGetRecordset
	dim cnnForumDC
	dim sRSSource

	sRSSource = "SELECT * FROM messages WHERE message_id=" & iActiveMessageId & ";"
	Call RunSQL(sRSSource, objRSGetRecordset)

	if not objRSGetRecordset.eof and not objRSGetRecordset.bof then	
		vTempRecords = objRSGetRecordset.GetRows
		' Reorganize Data
		lRows = UBound(vTempRecords, 1)
		lCols = UBound(vTempRecords, 2)
		ReDim vRecords(lCols)
		For lColIndex = 0 To lCols
		   ReDim vRow(lRows)
		   vRecords(lColIndex) = vRow
		Next
		For lColIndex = 0 To lCols
		   For lRowIndex = 0 To lRows
		      vRecords(lColIndex)(lRowIndex) = vTempRecords(lRowIndex, lColIndex)
		   Next
		Next
	end if
	objRSGetRecordset.close
	set objRSGetRecordset = nothing
	
	GetMessages=vRecords
end function

Function GetForums(iActiveForumID)
	Dim objRSGetRecordset
	dim cnnForumDC
	dim sRSSource

	sRSSource = "SELECT * FROM forums WHERE forum_id=" & iActiveForumId & ";"
	Call RunSQL(sRSSource, objRSGetRecordset)

	if not objRSGetRecordset.eof and not objRSGetRecordset.bof then	
		vTempRecords = objRSGetRecordset.GetRows		
		' Reorganize Data
		lRows = UBound(vTempRecords, 1)
		lCols = UBound(vTempRecords, 2)

		ReDim vRecords(lCols)
		For lColIndex = 0 To lCols
		   ReDim vRow(lRows)
		   vRecords(lColIndex) = vRow
		Next
		For lColIndex = 0 To lCols
		   For lRowIndex = 0 To lRows
		      vRecords(lColIndex)(lRowIndex) = vTempRecords(lRowIndex, lColIndex)
	      
		   Next
		Next
	end if
	objRSGetRecordset.close
	set objRSGetRecordset = nothing
	
	GetForums=vRecords
end function

Function GetThreads(ithreadid)
	Dim objRSGetRecordset
	dim cnnForumDC
	dim sRSSource
	
	sRSSource ="SELECT * FROM messages WHERE thread_id = (select thread_id from messages " & _
		"where message_id = " & ithreadid & ") ORDER BY message_id, thread_parent"
	Call RunSQL(sRSSource, objRSGetRecordset)

	if not objRSGetRecordset.eof and not objRSGetRecordset.bof then	
		vTempRecords = objRSGetRecordset.GetRows
		' Reorganize Data
		lRows = UBound(vTempRecords, 1)
		lCols = UBound(vTempRecords, 2)
		ReDim vRecords(lCols)
		For lColIndex = 0 To lCols
		   ReDim vRow(lRows)
		   vRecords(lColIndex) = vRow
		Next
		For lColIndex = 0 To lCols
		   For lRowIndex = 0 To lRows
		      vRecords(lColIndex)(lRowIndex) = vTempRecords(lRowIndex, lColIndex)
		   Next
		Next
	end if
	objRSGetRecordset.close
	set objRSGetRecordset = nothing

	GetThreads=vRecords	
end function

Function GetForumsBySect(ID)
	Dim objRSGetRecordset
	dim cnnForumDC
	dim sRSSource

	sRSSource = "SELECT * FROM forums where forum_section = "&ID&";"
	Call RunSQL(sRSSource, objRSGetRecordset)
	
	if not objRSGetRecordset.eof and not objRSGetRecordset.bof then	
		vTempRecords = objRSGetRecordset.GetRows
		' Reorganize Data
		lRows = UBound(vTempRecords, 1)
		lCols = UBound(vTempRecords, 2)
		ReDim vRecords(lCols)
		For lColIndex = 0 To lCols
		   ReDim vRow(lRows)
		   vRecords(lColIndex) = vRow
		Next
		For lColIndex = 0 To lCols
		   For lRowIndex = 0 To lRows
		      vRecords(lColIndex)(lRowIndex) = vTempRecords(lRowIndex, lColIndex)
		   Next
		Next
	end if
	objRSGetRecordset.close
	set objRSGetRecordset = nothing

	GetForumsBySect=vRecords
end function

Function GetAllForums()
	Dim objRSGetRecordset
	dim cnnForumDC
	dim sRSSource

	sRSSource = "SELECT * FROM forums;"
	Call RunSQL(sRSSource, objRSGetRecordset)
	
	if not objRSGetRecordset.eof and not objRSGetRecordset.bof then	
		vTempRecords = objRSGetRecordset.GetRows
		' Reorganize Data
		lRows = UBound(vTempRecords, 1)
		lCols = UBound(vTempRecords, 2)
		ReDim vRecords(lCols)
		For lColIndex = 0 To lCols
		   ReDim vRow(lRows)
		   vRecords(lColIndex) = vRow
		Next
		For lColIndex = 0 To lCols
		   For lRowIndex = 0 To lRows
		      vRecords(lColIndex)(lRowIndex) = vTempRecords(lRowIndex, lColIndex)
		   Next
		Next
	end if
	objRSGetRecordset.close
	set objRSGetRecordset = nothing

	GetAllForums=vRecords
end function

Function GetForumMessages(iActiveForumId)
	Dim objRSGetRecordset
	dim cnnForumDC
	dim sRSSource

	sRSSource = "SELECT * FROM messages WHERE forum_id=" & iActiveForumId & _
				" AND thread_parent=0 " & _
				" AND " & DB_DATEDELIMITER & MediumDate(dStartDate) &  DB_DATEDELIMITER & " <= message_timestamp " & _
				" AND message_timestamp <= " & DB_DATEDELIMITER & MediumDate(dEndDate) & DB_DATEDELIMITER & " " & _
				" ORDER BY thread_id DESC;"
	Call RunSQL(sRSSource, objRSGetRecordset)

	if not objRSGetRecordset.eof and not objRSGetRecordset.bof then
		vTempRecords = objRSGetRecordset.GetRows
		' Reorganize Data
		lRows = UBound(vTempRecords, 1)
		lCols = UBound(vTempRecords, 2)
		ReDim vRecords(lCols)
		For lColIndex = 0 To lCols
		ReDim vRow(lRows)
		vRecords(lColIndex) = vRow
		Next
		For lColIndex = 0 To lCols
		For lRowIndex = 0 To lRows
			vRecords(lColIndex)(lRowIndex) = vTempRecords(lRowIndex, lColIndex)
		Next
		Next
	end if
	objRSGetRecordset.close
	set objRSGetRecordset = nothing
	
	GetForumMessages=vRecords
end function

Function GetAllForumsMessageCounts()
	Dim objRSGetRecordset
	dim cnnForumDC
	dim sRSSource

	sRSSource = "SELECT forum_id, COUNT(*) " & _
				"FROM messages " & _
				"GROUP BY forum_id;"								
	Call RunSQL(sRSSource, objRSGetRecordset)

	if not objRSGetRecordset.eof and not objRSGetRecordset.bof then	
		vTempRecords = objRSGetRecordset.GetRows
		' Reorganize Data
		lRows = UBound(vTempRecords, 1)
		lCols = UBound(vTempRecords, 2)
		ReDim vRecords(lCols)
		For lColIndex = 0 To lCols
		   ReDim vRow(lRows)
		   vRecords(lColIndex) = vRow
		Next
		For lColIndex = 0 To lCols
		   For lRowIndex = 0 To lRows
		      vRecords(lColIndex)(lRowIndex) = vTempRecords(lRowIndex, lColIndex)
		   Next
		Next
	end if
	objRSGetRecordset.close
	set objRSGetRecordset = nothing
	
	GetAllForumsMessageCounts = vRecords
end function

Function GetMessageReplies(strThreadList)
	Dim objRSGetRecordset
	dim cnnForumDC
	dim sRSSource

	sRSSource = "SELECT thread_id, COUNT(*) " & _
				"FROM messages WHERE thread_id IN (" & strThreadList & ") " & _
				"GROUP BY thread_id ORDER BY thread_id DESC;"									
	Call RunSQL(sRSSource, objRSGetRecordset)

	if not objRSGetRecordset.eof and not objRSGetRecordset.bof then	
		vTempRecords = objRSGetRecordset.GetRows
		' Reorganize Data
		lRows = UBound(vTempRecords, 1)
		lCols = UBound(vTempRecords, 2)
		ReDim vRecords(lCols)
		For lColIndex = 0 To lCols
		   ReDim vRow(lRows)
		   vRecords(lColIndex) = vRow
		Next
		For lColIndex = 0 To lCols
		   For lRowIndex = 0 To lRows
		      vRecords(lColIndex)(lRowIndex) = vTempRecords(lRowIndex, lColIndex)
		   Next
		Next
	end if
	objRSGetRecordset.close
	set objRSGetRecordset = nothing

	GetMessageReplies = VRecords
end function

Function GetMessage(iActiveMessageId)
	Dim objRSGetRecordset
	dim cnnForumDC
	dim sRSSource

	sRSSource = "SELECT * FROM messages WHERE message_id=" & iActiveMessageId 											
	Call RunSQL(sRSSource, objRSGetRecordset)
	
	if not objRSGetRecordset.eof and not objRSGetRecordset.bof then	
		vTempRecords = objRSGetRecordset.GetRows
		' Reorganize Data
		lRows = UBound(vTempRecords, 1)
		lCols = UBound(vTempRecords, 2)
		ReDim vRecords(lCols)
		For lColIndex = 0 To lCols
		   ReDim vRow(lRows)
		   vRecords(lColIndex) = vRow
		Next
		For lColIndex = 0 To lCols
		   For lRowIndex = 0 To lRows
		      vRecords(lColIndex)(lRowIndex) = vTempRecords(lRowIndex, lColIndex)
		   Next
		Next
	end if
	objRSGetRecordset.close
	set objRSGetRecordset = nothing
	
	GetMessage = VRecords
end function

Function UpdateMessage(iActiveMessageId, strName, strEmail, strSubject, optnotify, iMessage )
	Dim objRSInsert
	Dim dTimeStamp
	Dim iNewMessageId
	dTimeStamp = Now()
	sSQL = "UPDATE Messages Set "
	sSQL = sSQL & "message_timestamp = " & DB_DATEDELIMITER & MediumDate(dTimeStamp) & DB_DATEDELIMITER & ", "
	sSQL = sSQL & "message_author = '" & strName & "', "
	sSQL = sSQL & "message_author_email = '" & strEmail & "', "
	sSQL = sSQL & "message_subject = '" & strSubject & "', "
	'sSQL = sSQL & "message_author_notify = '" & optnotify & "', " 
	sSQL = sSQL & "message_body ='" & iMessage & "', "
	sSQl = sSQL & "Message_Approved = 1 " 
	ssql = sSQL & "WHERE message_id = " & iActiveMessageId
	Call DoSQL(sSQL)	
end function


Function InsertRecord(forum_id, thread_id, thread_parent, thread_level, author, email, notify, subject, body)
	Dim objRSInsert
	Dim dTimeStamp
	Dim iNewMessageId
	
	dTimeStamp = Now()
	sSQL = "SELECT * FROM messages " & _
		    "WHERE message_timestamp=" & DB_DATEDELIMITER &  MediumDate(dTimeStamp) & DB_DATEDELIMITER & ";"
	Call RunSQL(sSQL, objRSInsert)
	objRSInsert.AddNew
	objRSInsert.Fields("message_timestamp") = MediumDate(dTimeStamp)
	objRSInsert.Fields("forum_id") = forum_id
	objRSInsert.Fields("thread_id") = thread_id
	objRSInsert.Fields("thread_parent") = thread_parent
	objRSInsert.Fields("thread_level") = thread_level
	objRSInsert.Fields("message_author") = author
	If email <> "" Then 
		objRSInsert.Fields("message_author_email") = email
	end if	
	objRSInsert.Fields("message_author_notify") = notify
	objRSInsert.Fields("message_subject") = subject
	objRSInsert.Fields("message_body") = body
	objRSInsert.Fields("Message_Approved") = 1
	
	objRSInsert.Update

    objRSInsert.Requery ' To be sure we have the message_id back from the DB.

	objRSInsert.MoveFirst

	iNewMessageId = objRSInsert.Fields("message_id")

	If thread_id = 0 Then
		objRSInsert.Fields("thread_id") = iNewMessageId
		objRSInsert.Update
	End If

	objRSInsert.Close
	Set objRSInsert = Nothing
	
	InsertRecord = iNewMessageId
End Function

Function InsertForum(iForumName, iForumDescription, iForumStart,su_section, iForumGroup)
	Dim objRSInsert
	Dim dTimeStamp
	Dim iNewMessageId
	
	dTimeStamp = Now()
	strSQL = "INSERT INTO forums "
		strSQL = strSQL & "(forum_name, "
		strSQL = strSQL & "forum_description, "
		strSQL = strSQL & "forum_start_date,forum_section, "
		strSQL = strSQL & "forum_grouping) " 
		strSQL = strSQL & "VALUES ("
		strSQL = strSQL & "'" & replace(iForumName,"'","''") & "',"
		strSQL = strSQL & "'" & replace(iForumDescription,"'","''") & "', "
		strSQL = strSQL & "'" & MediumDate(iForumStart) & "', "
		strSQL = strSQL & su_section & ", "
		strSQL = strSQL & "'" & iForumGroup & "'"
		strSQL = strSQL & ");"
	Call DoSQL(strSQL)
End Function

function GetSearch(strSearchType)
	Dim objRSGetRecordset
	dim cnnForumDC
	dim strSQL

	Select Case strSearchType
		Case "basic"
			strTemp = Replace(strKeyword, "'", "''")
			strSQL = "SELECT * FROM messages WHERE "
			strSQL = strSQL & "message_author LIKE '%" & strTemp & "%' OR "
			strSQL = strSQL & "message_author_email LIKE '%" & strTemp & "%' OR "
			strSQL = strSQL & "message_subject LIKE '%" & strTemp & "%' OR "
			strSQL = strSQL & "message_body LIKE '%" & strTemp & "%' "
			strSQL = strSQL & "ORDER BY message_timestamp DESC;"
		Case "advanced"
			strSQL = "SELECT * FROM messages WHERE "
			If strAuthor <> "" Then
				If Request.QueryString("a_type") = "contains" Then
					strSQL = strSQL & "message_author LIKE '%" & Replace(strAuthor, "'", "''") & "%' AND "
				Else
					strSQL = strSQL & "message_author = '" & Replace(strAuthor, "'", "''") & "' AND "
				End If
			End If
			If strEmail <> "" Then
				If Request.QueryString("e_type") = "contains" Then
					strSQL = strSQL & "message_author_email LIKE '%" & Replace(strEmail, "'", "''") & "%' AND "
				Else
					strSQL = strSQL & "message_author_email = '" & Replace(strEmail, "'", "''") & "' AND "
				End If
			End If
			If strSubject <> "" Then
				If Request.QueryString("s_type") = "contains" Then
					strSQL = strSQL & "message_subject LIKE '%" & Replace(strSubject, "'", "''") & "%' AND "
				Else
					strSQL = strSQL & "message_subject = '" & Replace(strSubject, "'", "''") & "' AND "
				End If
			End If
			If strBody <> "" Then
				If Request.QueryString("b_type") = "contains" Then
					strSQL = strSQL & "message_body LIKE '%" & Replace(strBody, "'", "''") & "%' AND "
				Else
					strTemp = Replace(strBody, "'", "''", 1, -1, 1)
					strTemp = Replace(strBody, ", ", " ", 1, -1, 1)					
					' Bad code ahead... alert alert!
					' I switch the strTemp var to an array.. so sue me!!!
					strTemp = Split(strTemp, " ", -1, 1)					
					For I = LBound(strTemp) To UBound(strTemp)
						strSQL = strSQL & "message_body LIKE '%" & strTemp(I) & "%' AND "
					Next 'I
				End If
			End If
			If Not IsNull(dStartDate) Then
				strSQL = strSQL & "message_timestamp >= " & DB_DATEDELIMITER & MediumDate(Replace(dStartDate, "'", "''")) & DB_DATEDELIMITER & " AND "
			End If
			If Not IsNull(dEndDate) Then
				strSQL = strSQL & "message_timestamp <= " & DB_DATEDELIMITER & MediumDate(Replace(dEndDate, "'", "''")) & DB_DATEDELIMITER & " AND "
			End If
			
			' remove the last "AND "
			strSQL = Left(strSQL, Len(strSQL) - 4)
			strSQL = strSQL & " ORDER BY message_timestamp DESC;"
		Case Else
			strSQL = ""
	End Select
	Call RunSQL(strSQL, objRSGetRecordset)

	if not objRSGetRecordset.eof and not objRSGetRecordset.bof then	
		vTempRecords = objRSGetRecordset.GetRows
		' Reorganize Data
		lRows = UBound(vTempRecords, 1)
		lCols = UBound(vTempRecords, 2)
		ReDim vRecords(lCols)
		For lColIndex = 0 To lCols
		   ReDim vRow(lRows)
		   vRecords(lColIndex) = vRow
		Next
		For lColIndex = 0 To lCols
		   For lRowIndex = 0 To lRows
		      vRecords(lColIndex)(lRowIndex) = vTempRecords(lRowIndex, lColIndex)
		   Next
		Next
	end if
	objRSGetRecordset.close
	set objRSGetRecordset = nothing	
	
	GetSearch = VRecords	
end function

Function UpdateForum(iForumID, iForumName, iForumDescription, su_section, iForumGroup)
	Dim objRSInsert
	Dim dTimeStamp
	Dim iNewMessageId

	strSQL = "UPDATE forums "
	strSQL = strSQL & " set forum_name ='" & replace(iForumName,"'","''") & "', "
	strSQL = strSQL & " forum_description = '" & replace(iForumDescription,"'","''") & "', "
	strSQL = strSQL & " forum_section = " & su_section & ", "
	strSQL = strSQL & " forum_grouping = '"  & replace(iForumGroup,"'","''") & "' "
	strSQL = strSQL & " WHERE Forum_ID = " & iForumID
	Call DoSQL(strSQL)
End Function

Function DeleteForum(iForumID)
	Dim objRSInsert
	Dim dTimeStamp
	Dim iNewMessageId
	
	strSQL = "DELETE From Messages "
	strSQL = strSQL & " WHERE Forum_ID = " & iForumID
	Call DoSQL(strSQL)

	strSQL = "DELETE From Forums "
	strSQL = strSQL & " WHERE forum_id = " & iForumID
	Call DoSQL(strSQL)
End Function

Function DeleteMessage(iMessageID)
	Dim objRSInsert
	Dim dTimeStamp
	Dim iNewMessageId
	
	strSQL = "DELETE FROM Messages WHERE Message_id = " & iMessageID & ""
	Call DoSQL(strSQL)
End Function
%>