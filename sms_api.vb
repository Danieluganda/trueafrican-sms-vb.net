'vb.net
Imports System.IO
Imports System.Net
Imports System.Text
Imports Newtonsoft.Json

Module SmsApi

    Sub Main()
        ' Example usage
        SendSms("256772123456", "Test SMS from True African")
    End Sub

    ' Function to send SMS
    Sub SendSms(msisdn As String, message As String)
        ' API Endpoint
        Dim url As String = "http://mysms.trueafrican.com/v1/api/esme/send"

        ' True African SMS credentials
        Dim username As String = "your_username"
        Dim password As String = "your_password"

        ' Create JSON payload
        Dim data As New Dictionary(Of String, Object) From {
            {"msisdn", New List(Of String) From {msisdn}},
            {"message", message},
            {"username", username},
            {"password", password}
        }

        Dim jsonData As String = JsonConvert.SerializeObject(data)

        ' Path to the log file
        Dim logFile As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sms_api.log")

        ' Log the request start time
        LogMessage(logFile, "Request Start Time: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
        LogMessage(logFile, "Sending SMS to: " & msisdn & " with message: " & message)

        ' Send the POST request
        Try
            Dim response As String = SendPostRequest(url, jsonData)

            ' Log the API response
            LogMessage(logFile, "API Response: " & response)

            ' Decode the response
            Dim responseData As Dictionary(Of String, Object) = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(response)

            ' Handle the response based on the code
            If responseData.ContainsKey("code") Then
                Dim responseCode As Integer = Convert.ToInt32(responseData("code"))

                Select Case responseCode
                    Case 200
                        LogMessage(logFile, "Message sent successfully.")
                    Case 204
                        LogMessage(logFile, "Request error: Check your request parameters.")
                    Case 209
                        LogMessage(logFile, "Authentication error: Invalid username or password.")
                    Case 207
                        LogMessage(logFile, "Bulk account error: Please check your account settings.")
                    Case Else
                        LogMessage(logFile, "Unexpected error: " & response)
                End Select
            Else
                LogMessage(logFile, "Unexpected error: " & response)
            End If

        Catch ex As Exception
            ' Log any error
            LogMessage(logFile, "Error: " & ex.Message)
        End Try

        ' Log the request end time
        LogMessage(logFile, "Request End Time: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
    End Sub

    ' Function to send a POST request
    Function SendPostRequest(url As String, jsonData As String) As String
        Dim request As HttpWebRequest = CType(WebRequest.Create(url), HttpWebRequest)
        request.Method = "POST"
        request.ContentType = "application/json"

        Dim data As Byte() = Encoding.UTF8.GetBytes(jsonData)
        request.ContentLength = data.Length

        Using stream = request.GetRequestStream()
            stream.Write(data, 0, data.Length)
        End Using

        Dim responseString As String
        Using response = CType(request.GetResponse(), HttpWebResponse)
            Using reader = New StreamReader(response.GetResponseStream())
                responseString = reader.ReadToEnd()
            End Using
        End Using

        Return responseString
    End Function

    ' Function to log messages to a file
    Sub LogMessage(filePath As String, message As String)
        Dim logMessage As String = "[" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "] " & message & Environment.NewLine
        File.AppendAllText(filePath, logMessage)
    End Sub

End Module


'vba-------------------------------

Sub sendSms(msisdn As String, message As String)
    ' API Endpoint
    Dim url As String
    url = "http://mysms.trueafrican.com/v1/api/esme/send"
    
    ' Your  credentials from as given by True African SMS
    Dim username As String
    Dim password As String
    username = "your_username"
    password = "your_password"
    
    ' Create JSON payload
    Dim jsonData As String
    jsonData = "{""msisdn"":[""" & msisdn & """],""message"":""" & message & """,""username"":""" & username & """,""password"":""" & password & """}"

    ' Path to the log file
    Dim logFile As String
    logFile = ThisWorkbook.Path & "\sms_api.log"

    ' Log the request
    Call logMessage(logFile, "Sending SMS to: " & msisdn & " with message: " & message)

    ' Create an XML HTTP request
    Dim http As Object
    On Error GoTo HttpError
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Initialize and send the POST request
    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json"
        .send jsonData
    End With

    ' Check the response status
    If http.Status = 200 Then
        ' Parse the response
        Dim jsonResponse As String
        jsonResponse = http.responseText

        ' Log the API response
        Call logMessage(logFile, "API Response: " & jsonResponse)

        ' Handle the response based on the code
        Dim responseCode As Long
        responseCode = GetResponseCode(jsonResponse)
        
        Select Case responseCode
            Case 200
                Call logMessage(logFile, "Message sent successfully.")
            Case 204
                Call logMessage(logFile, "Request error: Check your request parameters.")
            Case 209
                Call logMessage(logFile, "Authentication error: Invalid username or password.")
            Case 207
                Call logMessage(logFile, "Bulk account error: Please check your account settings.")
            Case Else
                Call logMessage(logFile, "Unexpected error: " & jsonResponse)
        End Select
    Else
        Call logMessage(logFile, "HTTP Error: " & http.Status)
    End If

    ' Clean up
    Set http = Nothing
    Exit Sub

HttpError:
    Call logMessage(logFile, "HTTP Error: " & Err.Description)
    Set http = Nothing
End Sub

' Function to log messages to a file
Sub logMessage(filePath As String, message As String)
    Dim fileNum As Integer
    On Error GoTo FileError

    fileNum = FreeFile
    Open filePath For Append As fileNum
    Print #fileNum, "[" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "] " & message
    Close fileNum
    Exit Sub

FileError:
    MsgBox "Error writing to log file: " & Err.Description
    Resume Next
End Sub

' Function to extract the response code from the JSON response
Function GetResponseCode(jsonResponse As String) As Long
    Dim codePos As Long
    Dim code As String

    codePos = InStr(jsonResponse, """code"":")
    If codePos > 0 Then
        code = Mid(jsonResponse, codePos + 7, 3)
        GetResponseCode = Val(code)
    Else
        GetResponseCode = 0
    End If
End Function

' Example usage
Sub testSendSms()
    Call sendSms("256701234567", "Test SMS from True African")
End Sub

