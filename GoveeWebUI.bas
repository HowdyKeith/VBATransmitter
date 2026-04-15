' --- Enhanced Web Interface ---
Public Function GenerateEnhancedGoveePage() As String
    On Error GoTo ErrorHandler
    
    Dim html As String
    html = "<!DOCTYPE html><html><head>"
    html = html & "<title>Enhanced Govee Control Center</title>"
    html = html & "<meta http-equiv='refresh' content='30'>"
    html = html & "<style>"
    html = html & "body { background: linear-gradient(135deg, #1a1a1a, #2d2d2d); color: #00ff88; font-family: 'Courier New', monospace; margin: 0; padding: 20px; }"
    html = html & ".container { max-width: 1400px; margin: 0 auto; }"
    html = html & ".header { background: rgba(0, 255, 136, 0.1); border: 1px solid #00ff88; padding: 20px; margin-bottom: 20px; border-radius: 10px; }"
    html = html & ".stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 20px; }"
    html = html & ".stat-card { background: rgba(0, 255, 136, 0.05); border: 1px solid #00ff88; padding: 15px; border-radius: 8px; text-align: center; }"
    html = html & ".stat-value { font-size: 24px; font-weight: bold; color: #00ff88; }"
    html = html & ".stat-label { font-size: 12px; color: #88ffaa; }"
    html = html & "table { width: 100%; border-collapse: collapse; margin-bottom: 20px; background: rgba(0, 255, 136, 0.02); }"
    html = html & "th, td { border: 1px solid #00ff88; padding: 8px; text-align: left; }"
    html = html & "th { background: rgba(0, 255, 136, 0.2); font-weight: bold; }"
    html = html & "tr:hover { background: rgba(0, 255, 136, 0.1); }"
    html = html & ".btn { background: #00ff88; color: #000; padding: 5px 10px; border: none; border-radius: 5px; text-decoration: none; margin: 2px; cursor: pointer; font-size: 11px; }"
    html = html & ".btn:hover { background: #00cc66; }"
    html = html & ".btn-danger { background: #ff4444; color: #fff; }"
    html = html & ".btn-preset { background: #4488ff; color: #fff; }"
    html = html & ".online { color: #00ff88; font-weight: bold; }"
    html = html & ".offline { color: #ff4444; font-weight: bold; }"
    html = html & ".section { background: rgba(0, 255, 136, 0.05); border: 1px solid #00ff88; margin-bottom: 20px; border-radius: 10px; }"
    html = html & ".section-header { background: rgba(0, 255, 136, 0.2); padding: 10px 15px; font-weight: bold; border-radius: 10px 10px 0 0; }"
    html = html & ".section-content { padding: 15px; }"
    html = html & ".tabs { display: flex; margin-bottom: 20px; }"
    html = html & ".tab { background: rgba(0, 255, 136, 0.1); border: 1px solid #00ff88; padding: 10px 20px; margin-right: 5px; cursor: pointer; border-radius: 5px 5px 0 0; }"
    html = html & ".tab.active { background: rgba(0, 255, 136, 0.3); }"
    html = html & ".input-group { margin: 10px 0; }"
    html = html & ".input-group label { display: inline-block; width: 120px; }"
    html = html & ".input-group input, .input-group select { background: #333; color: #00ff88; border: 1px solid #00ff88; padding: 5px; }"
    html = html & "</style>"
    html = html & "<script>"
    html = html & "function showTab(tabName) {"
    html = html & "  var tabs = document.getElementsByClassName('tab-content');"
    html = html & "  for(var i = 0; i < tabs.length; i++) { tabs[i].style.display = 'none'; }"
    html = html & "  document.getElementById(tabName).style.display = 'block';"
    html = html & "  var tabButtons = document.getElementsByClassName('tab');"
    html = html & "  for(var i = 0; i < tabButtons.length; i++) { tabButtons[i].classList.remove('active'); }"
    html = html & "  event.target.classList.add('active');"
    html = html & "}"
    html = html & "</script>"
    html = html & "</head><body>"
    
    html = html & "<div class='container'>"
    html = html & "<div class='header'>"
    html = html & "<h1>Enhanced Govee Control Center</h1>"
    html = html & "<p>Advanced IoT lighting control with scheduling, grouping, and automation</p>"
    html = html & "</div>"
    
    ' Statistics Overview
    Dim onlineCount As Long, totalDevices As Long, groupCount As Long, presetCount As Long, scheduleCount As Long
    totalDevices = deviceDict.count
    groupCount = groupDict.count
    presetCount = presetDict.count
    scheduleCount = scheduleDict.count
    
    Dim key As Variant
    For Each key In deviceDict.Keys
        Dim device As Object
        Set device = deviceDict(key)
        If device("isOnline") Then onlineCount = onlineCount + 1
    Next key
    
    html = html & "<div class='stats-grid'>"
    html = html & "<div class='stat-card'><div class='stat-value'>" & totalDevices & "</div><div class='stat-label'>Total Devices</div></div>"
    html = html & "<div class='stat-card'><div class='stat-value'>" & onlineCount & "</div><div class='stat-label'>Online Devices</div></div>"
    html = html & "<div class='stat-card'><div class='stat-value'>" & groupCount & "</div><div class='stat-label'>Device Groups</div></div>"
    html = html & "<div class='stat-card'><div class='stat-value'>" & presetCount & "</div><div class='stat-label'>Saved Presets</div></div>"
    html = html & "<div class='stat-card'><div class='stat-value'>" & scheduleCount & "</div><div class='stat-label'>Active Schedules</div></div>"
    html = html & "</div>"
    
    ' Tab Navigation
    html = html & "<div class='tabs'>"
    html = html & "<div class='tab active' onclick='showTab(""devices"")'>Devices</div>"
    html = html & "<div class='tab' onclick='showTab(""groups"")'>Groups</div>"
    html = html & "<div class='tab' onclick='showTab(""presets"")'>Presets</div>"
    html = html & "<div class='tab' onclick='showTab(""schedules"")'>Schedules</div>"
    html = html & "<div class='tab' onclick='showTab(""settings"")'>Settings</div>"
    html = html & "</div>"
    
    ' Devices Tab
    html = html & "<div id='devices' class='tab-content'>"
    html = html & "<div class='section'>"
    html = html & "<div class='section-header'>Device Management"
    html = html & "<a href='/govee/discover' class='btn' style='float: right;'>Discover Devices</a>"
    html = html & "</div>"
    html = html & "<div class='section-content'>"
    html = html & "<table>"
    html = html & "<tr><th>Name</th><th>MAC</th><th>IP</th><th>Model</th><th>Room</th><th>Group</th><th>Status</th><th>Power</th><th>Brightness</th><th>Color</th><th>Response</th><th>Actions</th></tr>"
    
    If totalDevices = 0 Then
        html = html & "<tr><td colspan='12'>No devices found. Click 'Discover Devices' to scan for Govee lights.</td></tr>"
    Else
        For Each key In deviceDict.Keys
            Set device = deviceDict(key)
            Dim statusColor As String, powerColor As String
            statusColor = IIf(device("isOnline"), "online", "offline")
            powerColor = IIf(device("power") = "On", "online", "")
            
            html = html & "<tr>"
            html = html & "<td><input type='text' value='" & device("name") & "' id='name_" & device("mac") & "' style='width:100px;'></td>"
            html = html & "<td>" & device("mac") & "</td>"
            html = html & "<td>" & device("ip") & "</td>"
            html = html & "<td>" & device("model") & "</td>"
            html = html & "<td><input type='text' value='" & device("room") & "' id='room_" & device("mac") & "' style='width:80px;'></td>"
            html = html & "<td>" & device("group") & "</td>"
            html = html & "<td class='" & statusColor & "'>" & IIf(device("isOnline"), "Online", "Offline") & "</td>"
            html = html & "<td class='" & powerColor & "'>" & device("power") & "</td>"
            html = html & "<td>" & device("brightness") & "%</td>"
            html = html & "<td style='background-color: rgb(" & device("color") & ");'>" & device("color") & "</td>"
            html = html & "<td>" & device("responseTime") & "ms</td>"
            html = html & "<td>"
            html = html & "<a href='/govee/device?mac=" & device("mac") & "&action=on' class='btn'>On</a>"
            html = html & "<a href='/govee/device?mac=" & device("mac") & "&action=off' class='btn btn-danger'>Off</a>"
            html = html & "<a href='/govee/device?mac=" & device("mac") & "&action=toggle' class='btn'>Toggle</a>"
            html = html & "<a href='/govee/device?mac=" & device("mac") & "&action=status' class='btn'>Status</a>"
            html = html & "</td>"
            html = html & "</tr>"
        Next key
    End If
    
    html = html & "</table>"
    html = html & "</div></div></div>"
    
    ' Groups Tab
    html = html & "<div id='groups' class='tab-content' style='display:none;'>"
    html = html & "<div class='section'>"
    html = html & "<div class='section-header'>Device Groups</div>"
    html = html & "<div class='section-content'>"
    
    ' Create Group Form
    html = html & "<h3>Create New Group</h3>"
    html = html & "<div class='input-group'><label>Group Name:</label><input type='text' id='groupName'></div>"
    html = html & "<div class='input-group'><label>Description:</label><input type='text' id='groupDesc'></div>"
    html = html & "<div class='input-group'><label>Devices:</label><select multiple id='groupDevices' style='height:100px;'>"
    
    For Each key In deviceDict.Keys
        Set device = deviceDict(key)
        html = html & "<option value='" & device("mac") & "'>" & device("name") & " (" & device("mac") & ")</option>"
    Next key
    
    html = html & "</select></div>"
    html = html & "<button class='btn' onclick='createGroup()'>Create Group</button>"
    
    ' Existing Groups
    html = html & "<h3>Existing Groups</h3>"
    html = html & "<table>"
    html = html & "<tr><th>Group Name</th><th>Description</th><th>Devices</th><th>Actions</th></tr>"
    
    If groupCount = 0 Then
        html = html & "<tr><td colspan='4'>No groups created yet.</td></tr>"
    Else
        For Each key In groupDict.Keys
            Dim group As Object
            Set group = groupDict(key)
            html = html & "<tr>"
            html = html & "<td>" & group("name") & "</td>"
            html = html & "<td>" & group("description") & "</td>"
            html = html & "<td>" & Replace(group("devices"), ",", ", ") & "</td>"
            html = html & "<td>"
            html = html & "<a href='/govee/group?name=" & group("name") & "&action=on' class='btn'>All On</a>"
            html = html & "<a href='/govee/group?name=" & group("name") & "&action=off' class='btn btn-danger'>All Off</a>"
            html = html & "</td>"
            html = html & "</tr>"
        Next key
    End If
    
    html = html & "</table>"
    html = html & "</div></div></div>"
    
    ' Presets Tab
    html = html & "<div id='presets' class='tab-content' style='display:none;'>"
    html = html & "<div class='section'>"
    html = html & "<div class='section-header'>Color Presets</div>"
    html = html & "<div class='section-content'>"
    
    ' Create Preset Form
    html = html & "<h3>Create New Preset</h3>"
    html = html & "<div class='input-group'><label>Preset Name:</label><input type='text' id='presetName'></div>"
    html = html & "<div class='input-group'><label>Description:</label><input type='text' id='presetDesc'></div>"
    html = html & "<div class='input-group'><label>Power:</label><select id='presetPower'><option>On</option><option>Off</option></select></div>"
    html = html & "<div class='input-group'><label>Brightness:</label><input type='range' id='presetBrightness' min='0' max='100' value='50'> <span id='brightnessValue'>50%</span></div>"
    html = html & "<div class='input-group'><label>Red:</label><input type='range' id='presetRed' min='0' max='255' value='255'> <span id='redValue'>255</span></div>"
    html = html & "<div class='input-group'><label>Green:</label><input type='range' id='presetGreen' min='0' max='255' value='255'> <span id='greenValue'>255</span></div>"
    html = html & "<div class='input-group'><label>Blue:</label><input type='range' id='presetBlue' min='0' max='255' value='255'> <span id='blueValue'>255</span></div>"
    html = html & "<button class='btn' onclick='createPreset()'>Create Preset</button>"
    
    ' Existing Presets
    html = html & "<h3>Saved Presets</h3>"
    html = html & "<table>"
    html = html & "<tr><th>Name</th><th>Description</th><th>Settings</th><th>Actions</th></tr>"
    
    If presetCount = 0 Then
        html = html & "<tr><td colspan='4'>No presets created yet.</td></tr>"
    Else
        For Each key In presetDict.Keys
            Dim preset As Object
            Set preset = presetDict(key)
            html = html & "<tr>"
            html = html & "<td>" & preset("name") & "</td>"
            html = html & "<td>" & preset("description") & "</td>"
            html = html & "<td>Power: " & preset("power") & ", Bright: " & preset("brightness") & "%, RGB(" & preset("red") & "," & preset("green") & "," & preset("blue") & ")</td>"
            html = html & "<td>"
            html = html & "<select id='presetTarget_" & key & "'>"
            html = html & "<option value=''>Select Device...</option>"
            For Each deviceKey In deviceDict.Keys
                Set device = deviceDict(deviceKey)
                html = html & "<option value='" & device("mac") & "'>" & device("name") & " (" & device("mac") & ")</option>"
            Next deviceKey
            html = html & "</select>"
            html = html & "<a href='#' class='btn btn-preset' onclick='applyPreset(""" & key & """)'>Apply</a>"
            html = html & "</td>"
            html = html & "</tr>"
        Next key
    End If
    
    html = html & "</table>"
    html = html & "</div></div></div>"
    
    ' Schedules Tab
    html = html & "<div id='schedules' class='tab-content' style='display:none;'>"
    html = html & "<div class='section'>"
    html = html & "<div class='section-header'>Automated Schedules</div>"
    html = html & "<div class='section-content'>"
    
    html = html & "<h3>Create Schedule</h3>"
    html = html & "<div class='input-group'><label>Schedule Name:</label><input type='text' id='scheduleName'></div>"
    html = html & "<div class='input-group'><label>Target:</label><select id='scheduleTarget'>"
    html = html & "<option value=''>Select Device or Group...</option>"
    For Each key In deviceDict.Keys
        Set device = deviceDict(key)
        html = html & "<option value='" & device("mac") & "'>Device: " & device("name") & "</option>"
    Next key
    For Each key In groupDict.Keys
        Set group = groupDict(key)
        html = html & "<option value='" & group("name") & "'>Group: " & group("name") & "</option>"
    Next key
    html = html & "</select></div>"
    html = html & "<div class='input-group'><label>Action:</label><select id='scheduleAction'>"
    html = html & "<option value='on'>Turn On</option><option value='off'>Turn Off</option>"
    For Each key In presetDict.Keys
        html = html & "<option value='" & key & "'>Apply Preset: " & key & "</option>"
    Next key
    html = html & "</select></div>"
    html = html & "<div class='input-group'><label>Time:</label><input type='time' id='scheduleTime'></div>"
    html = html & "<div class='input-group'><label>Days:</label>"
    html = html & "<input type='checkbox' value='Mon'> Mon <input type='checkbox' value='Tue'> Tue <input type='checkbox' value='Wed'> Wed "
    html = html & "<input type='checkbox' value='Thu'> Thu <input type='checkbox' value='Fri'> Fri <input type='checkbox' value='Sat'> Sat <input type='checkbox' value='Sun'> Sun"
    html = html & "</div>"
    html = html & "<button class='btn' onclick='createSchedule()'>Create Schedule</button>"
    
    html = html & "<h3>Active Schedules</h3>"
    html = html & "<table>"
    html = html & "<tr><th>Name</th><th>Target</th><th>Action</th><th>Time</th><th>Days</th><th>Status</th><th>Last Run</th><th>Actions</th></tr>"
    
    If scheduleCount = 0 Then
        html = html & "<tr><td colspan='8'>No schedules created yet.</td></tr>"
    Else
        For Each key In scheduleDict.Keys
            Dim schedule As Object
            Set schedule = scheduleDict(key)
            html = html & "<tr>"
            html = html & "<td>" & schedule("name") & "</td>"
            html = html & "<td>" & schedule("target") & "</td>"
            html = html & "<td>" & schedule("action") & "</td>"
            html = html & "<td>" & schedule("time") & "</td>"
            html = html & "<td>" & schedule("days") & "</td>"
            html = html & "<td>" & IIf(schedule("enabled"), "Enabled", "Disabled") & "</td>"
            html = html & "<td>" & format(schedule("lastRun"), "yyyy-mm-dd hh:mm") & "</td>"
            html = html & "<td>"
            html = html & "<a href='/govee/schedule?name=" & schedule("name") & "&action=toggle' class='btn'>Toggle</a>"
            html = html & "<a href='/govee/schedule?name=" & schedule("name") & "&action=delete' class='btn btn-danger'>Delete</a>"
            html = html & "</td>"
            html = html & "</tr>"
        Next key
    End If
    
    html = html & "</table>"
    html = html & "</div></div></div>"
    
    ' Settings Tab
    html = html & "<div id='settings' class='tab-content' style='display:none;'>"
    html = html & "<div class='section'>"
    html = html & "<div class='section-header'>System Settings</div>"
    html = html & "<div class='section-content'>"
    html = html & "<h3>Monitoring Settings</h3>"
    html = html & "<p>Auto-refresh interval: 5 minutes</p>"
    html = html & "<p>Schedule check interval: 1 minute</p>"
    html = html & "<p>Device timeout: 10 minutes</p>"
    html = html & "<h3>Network Settings</h3>"
    html = html & "<p>Discovery port: " & DISCOVERY_PORT & "</p>"
    html = html & "<p>Control port: " & CONTROL_PORT & "</p>"
    html = html & "<p>Local bind port: " & LOCAL_BIND_PORT & "</p>"
    html = html & "<h3>Data Management</h3>"
    html = html & "<a href='/govee/export/devices' class='btn'>Export Device List</a>"
    html = html & "<a href='/govee/export/presets' class='btn'>Export Presets</a>"
    html = html & "<a href='/govee/export/schedules' class='btn'>Export Schedules</a>"
    html = html & "<a href='/govee/reset' class='btn btn-danger' onclick='return confirm(\"Reset all data?\")'>Reset All Data</a>"
    html = html & "</div></div></div>"
    
    html = html & "</div>"
    html = html & "</body></html>"
    
    GenerateEnhancedGoveePage = html
    Exit Function
    
ErrorHandler:
    DebuggingLog.DebugLog "Error generating enhanced Govee page: " & Err.description
    GenerateEnhancedGoveePage = "<html><body><h1>Error generating page</h1><p>" & Err.description & "</p></body></html>"
End Function

