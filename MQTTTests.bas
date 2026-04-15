
Option Explicit

'***************************************************************
' MQTT Demo Launcher and Test Suite
' Purpose: Complete demo to test MQTT integration with TrafficManager
'***************************************************************

' Main demo launcher - call this to start everything
Public Sub LaunchMQTTDemo()
    On Error GoTo ErrorHandler
    
    Debug.Print "========================================="
    Debug.Print "   VBA MQTT Demo - Starting Up"
    Debug.Print "========================================="
    
    ' Initialize debugging if not already done
    DebuggingLog.DebugLog "[DEMO] Starting MQTT Demo Suite"
    
    ' Start TrafficManager with all servers including MQTT
    TrafficManager.InitializeTrafficManager
    TrafficManager.RunAllServers
    
    ' Wait a moment for everything to initialize
    Application.Wait Now + TimeValue("00:00:03")
    
    ' Run MQTT-specific tests
    RunMQTTTests
    
    ' Display status
    ShowMQTTStatus
    
    Debug.Print "========================================="
    Debug.Print "   MQTT Demo Started Successfully!"
    Debug.Print "   Visit http://localhost:8080/mqtt"
    Debug.Print "   Press ESC in Excel to stop servers"
    Debug.Print "========================================="
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "[DEMO ERROR] " & Err.description
    DebuggingLog.DebugLog "[DEMO ERROR] " & Err.description
End Sub

' Run comprehensive MQTT tests
Private Sub RunMQTTTests()
    DebuggingLog.DebugLog "[DEMO] Running MQTT tests..."
    
    ' Test 1: Basic Publishing
    If IsConnected() Then
        PublishMQTTAdvanced "demo/test1", "Hello World from VBA!", MQTT_QOS0
        PublishMQTTAdvanced "demo/test2", "QoS 1 message", MQTT_QOS1
        PublishMQTTAdvanced "demo/test3", "QoS 2 message", MQTT_QOS2
        DebuggingLog.DebugLog "[DEMO] Published 3 test messages"
    End If
    
    ' Test 2: Subscribe to various topics
    SubscribeMQTTAdvanced "demo/#", MQTT_QOS1
    SubscribeMQTTAdvanced "sensors/+/data", MQTT_QOS0
    SubscribeMQTTAdvanced "alerts/critical", MQTT_QOS2
    SubscribeMQTTAdvanced "system/heartbeat", MQTT_QOS0
    DebuggingLog.DebugLog "[DEMO] Subscribed to 4 test topics"
    
    ' Test 3: Simulated sensor data
    Dim i As Integer
    For i = 1 To 5
        PublishMQTTAdvanced "sensors/temperature", CStr(18 + Rnd() * 12), MQTT_QOS0
        PublishMQTTAdvanced "sensors/humidity", CStr(40 + Rnd() * 30), MQTT_QOS0
        PublishMQTTAdvanced "sensors/pressure", CStr(1010 + Rnd() * 20), MQTT_QOS1
        Application.Wait Now + TimeValue("00:00:01")
    Next i
    DebuggingLog.DebugLog "[DEMO] Published 15 simulated sensor readings"
    
    ' Test 4: System status messages
    PublishMQTTAdvanced "system/status", "VBA MQTT System Online", MQTT_QOS1
    PublishMQTTAdvanced "system/memory", "Memory usage: " & CStr(Int(Rnd() * 1000)) & "MB", MQTT_QOS0
    PublishMQTTAdvanced "system/cpu", "CPU usage: " & CStr(Int(Rnd() * 100)) & "%", MQTT_QOS0
    DebuggingLog.DebugLog "[DEMO] Published system status messages"
    
    ' Test 5: Alert simulation
    PublishMQTTAdvanced "alerts/info", "Demo alert: All systems operational", MQTT_QOS0
    PublishMQTTAdvanced "alerts/warning", "Demo warning: High CPU usage detected", MQTT_QOS1
    PublishMQTTAdvanced "alerts/critical", "Demo critical: Emergency test message", MQTT_QOS2
    DebuggingLog.DebugLog "[DEMO] Published alert simulation messages"
End Sub

' Display current MQTT status
Private Sub ShowMQTTStatus()
    Debug.Print ""
    Debug.Print "--- MQTT Status Report ---"
    Debug.Print "Connected: " & IIf(IsConnected(), "YES", "NO")
    If IsConnected() Then
        Debug.Print "Broker: " & mqttClientData.brokerAddress & ":" & mqttClientData.brokerPort
        Debug.Print "Client ID: " & mqttClientData.clientID
        Debug.Print "Subscriptions: " & mqttSubscriptions.count
        Debug.Print "Messages: " & mqttMessageLog.count
        Debug.Print "Pending ACKs: " & mqttPendingAcks.count
    End If
    Debug.Print "--- End Status Report ---"
    Debug.Print ""
End Sub

' Quick test for individual MQTT operations
Public Sub QuickMQTTTest()
    If Not IsConnected() Then
        Debug.Print "[TEST] MQTT not connected - starting connection..."
        InitializeMQTTAdvanced "test.mosquitto.org", 1883
        ConnectMQTTAdvanced
        Application.Wait Now + TimeValue("00:00:02")
    End If
    
    If IsConnected() Then
        PublishMQTTAdvanced "vba/quicktest", "Quick test at " & format(Now, "hh:mm:ss"), MQTT_QOS1
        SubscribeMQTTAdvanced "vba/response", MQTT_QOS1
        Debug.Print "[TEST] Quick MQTT test completed"
    Else
        Debug.Print "[TEST] Failed to connect to MQTT broker"
    End If
End Sub

' Stress test for MQTT
Public Sub MQTTStressTest()
    If Not IsConnected() Then
        Debug.Print "[STRESS] MQTT not connected - cannot run stress test"
        Exit Sub
    End If
    
    Debug.Print "[STRESS] Starting MQTT stress test..."
    Dim startTime As Double
    startTime = Timer
    
    ' Publish 100 messages rapidly
    Dim i As Integer
    For i = 1 To 100
        PublishMQTTAdvanced "stress/test" & (i Mod 10), "Message #" & i & " at " & format(Now, "hh:mm:ss.000"), MQTT_QOS0
        If i Mod 10 = 0 Then
            DoEvents  ' Allow other processes
            Application.Wait Now + TimeValue("00:00:00.1")  ' 100ms delay every 10 messages
        End If
    Next i
    
    Dim endTime As Double
    endTime = Timer
    Debug.Print "[STRESS] Published 100 messages in " & format(endTime - startTime, "0.00") & " seconds"
    Debug.Print "[STRESS] Rate: " & format(100 / (endTime - startTime), "0.0") & " messages/second"
End Sub

' MQTT monitoring function
Public Sub MonitorMQTTActivity()
    Debug.Print "========================================="
    Debug.Print "   MQTT Activity Monitor"
    Debug.Print "========================================="
    
    If Not IsConnected() Then
        Debug.Print "MQTT is not connected."
        Exit Sub
    End If
    
    ' Display recent activity
    If mqttMessageLog.count > 0 Then
        Debug.Print "Recent MQTT Messages:"
        Dim i As Integer
        Dim startIdx As Integer
        startIdx = IIf(mqttMessageLog.count > 20, mqttMessageLog.count - 19, 1)
        
        For i = startIdx To mqttMessageLog.count
            Debug.Print "  " & mqttMessageLog.item(i)
        Next i
    Else
        Debug.Print "No MQTT messages logged yet."
    End If
    
    Debug.Print ""
    Debug.Print "Statistics:"
    Debug.Print "  Total Messages: " & mqttMessageLog.count
    Debug.Print "  Active Subscriptions: " & mqttSubscriptions.count
    Debug.Print "  Pending Acknowledgments: " & mqttPendingAcks.count
    Debug.Print "  Reconnection Attempts: " & mqttClientData.reconnectAttempts
    Debug.Print ""
End Sub

' Stop MQTT demo
Public Sub StopMQTTDemo()
    Debug.Print "========================================="
    Debug.Print "   Stopping MQTT Demo"
    Debug.Print "========================================="
    
    TrafficManager.StopAllServers
    
    Debug.Print "MQTT Demo stopped."
End Sub

' Interactive MQTT publisher (for testing from immediate window)
'Public Sub PublishTestMessage(Optional topic As String = "vba/test", Optional message As String = "")
 '   If message = "" Then
  '      message =
