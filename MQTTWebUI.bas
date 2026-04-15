'---------------------------------------------------------------
' MQTT Web Interface Handler (add to your HttpServer module)
'---------------------------------------------------------------

' Add this function to handle MQTT-related web requests in your HttpServer module:
Public Function HandleMQTTWebRequest(ByVal requestPath As String, ByVal queryString As String) As String
    Dim html As String
    
    Select Case LCase(requestPath)
        Case "/mqtt", "/mqtt/"
            html = GetMQTTDashboardHTML()
        Case "/mqtt/log"
            html = GetMQTTLogHTMLAdvanced()
        Case "/mqtt/publish"
            html = HandleMQTTPublishWeb(queryString)
        Case "/mqtt/subscribe"
            html = HandleMQTTSubscribeWeb(queryString)
        Case "/mqtt/status"
            html = GetMQTTStatusJSON()
        Case "/mqtt/test"
            html = HandleMQTTTestWeb()
        Case Else
            html = GetMQTT404HTML()
    End Select
    
    HandleMQTTWebRequest = html
End Function

' MQTT Dashboard HTML
Private Function GetMQTTDashboardHTML() As String
    Dim html As String
    html = "<!DOCTYPE html><html><head><title>VBA MQTT Dashboard</title>"
    html = html & "<meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>"
    html = html & "<style>"
    html = html & "body{font-family:Arial,sans-serif;margin:0;padding:20px;background:#f5f5f5;}"
    html = html & ".container{max-width:1200px;margin:0 auto;}"
    html = html & ".card{background:white;border-radius:8px;padding:20px;margin:10px 0;box-shadow:0 2px 4px rgba(0,0,0,0.1);}"
    html = html & ".status{padding:10px;border-radius:4px;margin:10px 0;}"
    html = html & ".connected{background:#d4edda;color:#155724;border:1px solid #c3e6cb;}"
    html = html & ".disconnected{background:#f8d7da;color:#721c24;border:1px solid #f5c6cb;}"
    html = html & ".button{background:#007bff;color:white;padding:10px 20px;border:none;border-radius:4px;cursor:pointer;margin:5px;}"
    html = html & ".button:hover{background:#0056b3;}"
    html = html & ".input-group{margin:10px 0;}"
    html = html & ".input-group label{display:block;margin-bottom:5px;font-weight:bold;}"
    html = html & ".input-group input,.input-group select{width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;}"
    html = html & ".log{background:#f8f9fa;border:1px solid #e9ecef;padding:10px;max-height:300px;overflow-y:auto;font-family:monospace;font-size:12px;}"
    html = html & ".stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:10px;}"
    html = html & ".stat-item{text-align:center;}"
    html = html & ".stat-value{font-size:24px;font-weight:bold;color:#007bff;}"
    html = html & ".refresh-btn{float:right;}"
    html = html & "</style>"
    html = html & "<script>"
    html = html & "function refreshPage(){window.location.reload();}"
    html = html & "function publishMessage(){var topic=document.getElementById('pubTopic').value;var msg=document.getElementById('pubMessage').value;var qos=document.getElementById('pubQos').value;window.location.href='/mqtt/publish?topic='+encodeURIComponent(topic)+'&message='+encodeURIComponent(msg)+'&qos='+qos;}"
    html = html & "function subscribeToTopic(){var topic=document.getElementById('subTopic').value;var qos=document.getElementById('subQos').value;window.location.href='/mqtt/subscribe?topic='+encodeURIComponent(topic)+'&qos='+qos;}"
    html = html & "function runTest(){window.location.href='/mqtt/test';}"
    html = html & "setInterval(function(){fetch('/mqtt/status').then(r=>r.json()).then(data=>{document.getElementById('statusInfo').innerHTML='<strong>Status:</strong> '+(data.connected?'<span class=\"connected\">CONNECTED</span>':'<span class=\"disconnected\">DISCONNECTED</span>')+'<br><strong>Messages:</strong> '+data.messages+'<br><strong>Subscriptions:</strong> '+data.subscriptions;});},5000);"
    html = html & "</script>"
    html = html & "</head><body>"
    
    html = html & "<div class='container'>"
    html = html & "<h1>VBA MQTT Dashboard</h1>"
    html = html & "<button class='button refresh-btn' onclick='refreshPage()'>Refresh</button>"
    
    ' Status Card
    html = html & "<div class='card'>"
    html = html & "<h2>Connection Status</h2>"
    html = html & "<div id='statusInfo'>"
    If mqttEnabled And IsConnected() Then
        html = html & "<div class='status connected'>CONNECTED to " & mqttClientData.brokerAddress & ":" & mqttClientData.brokerPort & "</div>"
        html = html & "<p><strong>Client ID:</strong> " & mqttClientData.clientID & "</p>"
    Else
        html = html & "<div class='status disconnected'>DISCONNECTED</div>"
    End If
    html = html & "</div>"
    html = html & "</div>"
    
    ' Statistics Card
    html = html & "<div class='card'>"
    html = html & "<h2>Statistics</h2>"
    html = html & "<div class='stats'>"
    html = html & "<div class='stat-item'><div class='stat-value'>" & IIf(mqttSubscriptions Is Nothing, 0, mqttSubscriptions.count) & "</div><div>Subscriptions</div></div>"
    html = html & "<div class='stat-item'><div class='stat-value'>" & IIf(mqttMessageLog Is Nothing, 0, mqttMessageLog.count) & "</div><div>Messages</div></div>"
    html = html & "<div class='stat-item'><div class='stat-value'>" & IIf(mqttPendingAcks Is Nothing, 0, mqttPendingAcks.count) & "</div><div>Pending ACKs</div></div>"
    html = html & "<div class='stat-item'><div class='stat-value'>" & mqttClientData.reconnectAttempts & "</div><div>Reconnect Attempts</div></div>"
    html = html & "</div>"
    html = html & "</div>"
    
    ' Publish Card
    html = html & "<div class='card'>"
    html = html & "<h2>Publish Message</h2>"
    html = html & "<div class='input-group'><label>Topic:</label><input type='text' id='pubTopic' placeholder='sensors/temperature' value='vba/test'></div>"
    html = html & "<div class='input-group'><label>Message:</label><input type='text' id='pubMessage' placeholder='Your message here' value='Hello from VBA Dashboard!'></div>"
    html = html & "<div class='input-group'><label>QoS:</label><select id='pubQos'><option value='0'>QoS 0</option><option value='1' selected>QoS 1</option><option value='2'>QoS 2</option></select></div>"
    html = html & "<button class='button' onclick='publishMessage()'>Publish</button>"
    html = html & "</div>"
    
    ' Subscribe Card
    html = html & "<div class='card'>"
    html = html & "<h2>Subscribe to Topic</h2>"
    html = html & "<div class='input-group'><label>Topic:</label><input type='text' id='subTopic' placeholder='sensors/+' value='test/#'></div>"
    html = html & "<div class='input-group'><label>QoS:</label><select id='subQos'><option value='0'>QoS 0</option><option value='1' selected>QoS 1</option><option value='2'>QoS 2</option></select></div>"
    html = html & "<button class='button' onclick='subscribeToTopic()'>Subscribe</button>"
    html = html & "</div>"
    
    ' Test Card
    html = html & "<div class='card'>"
    html = html & "<h2>Test Functions</h2>"
    html = html & "<button class='button' onclick='runTest()'>Run MQTT Test</button>"
    html = html & "<button class='button' onclick='window.open(\" / MQTT / Log \ ")'>View Full Log</button>"
    html = html & "</div>"
    
    ' Recent Messages Card
    html = html & "<div class='card'>"
    html = html & "<h2>Recent Messages</h2>"
    html = html & "<div class='log'>"
    If Not mqttMessageLog Is Nothing And mqttMessageLog.count > 0 Then
        Dim i As Long
        For i = IIf(mqttMessageLog.count > 10, mqttMessageLog.count - 9, 1) To mqttMessageLog.count
            html = html & mqttMessageLog.item(i) & "<br>"
        Next i
    Else
        html = html & "No messages yet..."
    End If
    html = html & "</div>"
    html = html & "</div>"
    
    html = html & "</div></body></html>"
    GetMQTTDashboardHTML = html
End Function

' Handle MQTT publish via web
Private Function HandleMQTTPublishWeb(ByVal queryString As String) As String
    Dim topic As String, message As String, qos As Integer
    topic = GetQueryParam(queryString, "topic")
    message = GetQueryParam(queryString, "message")
    qos = CInt(IIf(GetQueryParam(queryString, "qos") = "", "0", GetQueryParam(queryString, "qos")))
    
    If Len(topic) > 0 And Len(message) > 0 Then
        If mqttEnabled And IsConnected() Then
            PublishMQTTAdvanced topic, message, qos
            HandleMQTTPublishWeb = "<html><head><meta http-equiv='refresh' content='2;url=/mqtt'></head><body><h2>Message Published!</h2><p>Topic: " & topic & "</p><p>Message: " & message & "</p><p>QoS: " & qos & "</p><p>Redirecting...</p></body></html>"
        Else
            HandleMQTTPublishWeb = "<html><head><meta http-equiv='refresh' content='3;url=/mqtt'></head><body><h2>Error: MQTT Not Connected</h2><p>Cannot publish message. MQTT client is not connected.</p><p>Redirecting...</p></body></html>"
        End If
    Else
        HandleMQTTPublishWeb = "<html><head><meta http-equiv='refresh' content='3;url=/mqtt'></head><body><h2>Error: Missing Parameters</h2><p>Topic and message are required.</p><p>Redirecting...</p></body></html>"
    End If
End Function

' Handle MQTT subscribe via web
Private Function HandleMQTTSubscribeWeb(ByVal queryString As String) As String
    Dim topic As String, qos As Integer
    topic = GetQueryParam(queryString, "topic")
    qos = CInt(IIf(GetQueryParam(queryString, "qos") = "", "0", GetQueryParam(queryString, "qos")))
    
    If Len(topic) > 0 Then
        If mqttEnabled And IsConnected() Then
            SubscribeMQTTAdvanced topic, qos
            HandleMQTTSubscribeWeb = "<html><head><meta http-equiv='refresh' content='2;url=/mqtt'></head><body><h2>Subscribed!</h2><p>Topic: " & topic & "</p><p>QoS: " & qos & "</p><p>Redirecting...</p></body></html>"
        Else
            HandleMQTTSubscribeWeb = "<html><head><meta http-equiv='refresh' content='3;url=/mqtt'></head><body><h2>Error: MQTT Not Connected</h2><p>Cannot subscribe. MQTT client is not connected.</p><p>Redirecting...</p></body></html>"
        End If
    Else
        HandleMQTTSubscribeWeb = "<html><head><meta http-equiv='refresh' content='3;url=/mqtt'></head><body><h2>Error: Missing Topic</h2><p>Topic is required for subscription.</p><p>Redirecting...</p></body></html>"
    End If
End Function

' Handle MQTT test via web
Private Function HandleMQTTTestWeb() As String
    If mqttEnabled And IsConnected() Then
        ' Run test messages
        PublishMQTTAdvanced "vba/test/web", "Web test message at " & format(Now, "hh:mm:ss"), MQTT_QOS1
        PublishMQTTAdvanced "sensors/test/temperature", CStr(20 + Rnd() * 15), MQTT_QOS0
        PublishMQTTAdvanced "sensors/test/humidity", CStr(45 + Rnd() * 20), MQTT_QOS0
        PublishMQTTAdvanced "system/test/status", "Test completed successfully", MQTT_QOS1
        
        HandleMQTTTestWeb = "<html><head><meta http-equiv='refresh' content='3;url=/mqtt'></head><body><h2>Test Messages Sent!</h2><p>4 test messages have been published to various topics.</p><p>Check the message log to see them.</p><p>Redirecting...</p></body></html>"
    Else
        HandleMQTTTestWeb = "<html><head><meta http-equiv='refresh' content='3;url=/mqtt'></head><body><h2>Error: MQTT Not Connected</h2><p>Cannot send test messages. MQTT client is not connected.</p><p>Redirecting...</p></body></html>"
    End If
End Function

' Get MQTT status as JSON
Private Function GetMQTTStatusJSON() As String
    GetMQTTStatusJSON = GetMQTTStatsJSONAdvanced()
End Function

' 404 for MQTT paths
Private Function GetMQTT404HTML() As String
    GetMQTT404HTML = "<html><head><title>MQTT - Not Found</title></head><body><h1>404 - MQTT Page Not Found</h1><p><a href='/mqtt'>Return to MQTT Dashboard</a></p></body></html>"
End Function

' Helper function to extract query parameters (add to HttpServer if not already present)
Private Function GetQueryParam(ByVal queryString As String, ByVal paramName As String) As String
    Dim params() As String
    Dim i As Integer
    Dim keyValue() As String
    
    If Len(queryString) = 0 Then
        GetQueryParam = ""
        Exit Function
    End If
    
    params = Split(queryString, "&")
    For i = 0 To UBound(params)
        keyValue = Split(params(i), "=")
        If UBound(keyValue) >= 1 Then
            If LCase(keyValue(0)) = LCase(paramName) Then
                GetQueryParam = Replace(keyValue(1), "+", " ")  ' Basic URL decode
                GetQueryParam = Replace(GetQueryParam, "%20", " ")
                Exit Function
            End If
        End If
    Next i
    
    GetQueryParam = ""
End Function
