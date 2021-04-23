' Copyright 2021 clyvari https://github.com/clyvari
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

Private Sub AutoRSVP(oMeetingItem_Request As MeetingItem, bAcceptMeeting As Boolean, bSendResponse As Boolean, Optional bDebug As Boolean = False)
    
    Dim strAutoRSVP_Title As String: strAutoRSVP_Title = "AutoRSVP - Debug"
    
    Dim strMessageClass As String: strMessageClass = "IPM.Schedule.Meeting.Request"
    If oMeetingItem_Request.MessageClass <> strMessageClass Then
        If bDebug Then
            MsgBox "MessageClass should be " + strMessageClass + " not " + oMeetingItem_Request.MessageClass, vbInformation, strAutoRSVP_Title
        End If
      Exit Sub
    End If
     
    Dim oAppointmentItem As AppointmentItem
    Set oAppointmentItem = oMeetingItem_Request.GetAssociatedAppointment(True)
    
    
    Dim strRequestAction As String
    If bAcceptMeeting Then
        strRequestAction = "accepted"
    Else
        strRequestAction = "declined"
    End If
    
    Dim strMeetingInfo As String
    strMeetingInfo = "Meeting request" + strRequestAction + ":" + vbNewLine
    strMeetingInfo = strMeetingInfo + oAppointmentItem.Subject + "(" + oMeetingItem_Request.SenderName + ") @ " + Format(oAppointmentItem.Start)
    
    
    Dim oMeetingResponse As OlMeetingResponse
    If bAcceptMeeting Then
        oMeetingResponse = olMeetingAccepted
    Else
        oMeetingResponse = olMeetingDeclined
    End If
    
    
    Dim oMeetingItem_Response As MeetingItem
    Set oMeetingItem_Response = oAppointmentItem.Respond(oMeetingResponse, True)
    
    If bSendResponse Then
        oMeetingItem_Response.Save
        oMeetingItem_Response.Send
    Else
        oMeetingItem_Response.Close (olDiscard)
        oMeetingItem_Response.Delete
    End If
    
    If bDebug Then
        MsgBox strMeetingInfo, vbInformation, strAutoRSVP_Title
    End If
    
    oMeetingItem_Request.Delete
    
End Sub

Public Sub AutoRSVP_DeclineSilently(oMeetingItem_Request As MeetingItem)
  AutoRSVP oMeetingItem_Request, False, False
End Sub
Public Sub AutoRSVP_AcceptSilently(oMeetingItem_Request As MeetingItem)
  AutoRSVP oMeetingItem_Request, True, False
End Sub
Public Sub AutoRSVP_DeclineAndSend(oMeetingItem_Request As MeetingItem)
  AutoRSVP oMeetingItem_Request, False, True
End Sub
Public Sub AutoRSVP_AcceptAndSend(oMeetingItem_Request As MeetingItem)
  AutoRSVP oMeetingItem_Request, True, True
End Sub
