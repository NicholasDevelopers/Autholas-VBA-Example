# Autholas VBA Excel Authentication System

A comprehensive VBA implementation for Excel that integrates with Autholas authentication service, featuring hardware ID verification and user session management.

## Features

- User authentication via Autholas API
- Hardware ID generation using computer name
- Comprehensive error handling with user-friendly messages
- Session management across UserForms
- Real-time API communication
- Secure credential validation
- Multi-device tracking support

## Prerequisites

- **Microsoft Excel 2010** or later (2013/2016/2019/365 recommended)
- **Windows OS** (Windows 7 or later)
- **Internet connection** for API communication
- **VBA-JSON Library** (VBA-Tools JsonConverter)

## Installation Guide

### Step 1: Download VBA-JSON Library

The project requires the VBA-JSON library for JSON parsing and serialization.

#### Download from GitHub Release:

1. Visit the official VBA-JSON release page:
   ```
   https://github.com/VBA-tools/VBA-JSON/releases
   ```

2. Download the latest release file:
   - **JsonConverter.bas** (recommended for this project)

3. Save the file to your computer

### Step 2: Create Excel Project

1. Open Microsoft Excel
2. Press `Alt + F11` to open VBA Editor
3. Create a new Excel Workbook or open existing one
4. Save as **Excel Macro-Enabled Workbook** (.xlsm)

### Step 3: Import VBA-JSON Library

1. In VBA Editor, go to **File** → **Import File**
2. Browse and select the downloaded **JsonConverter.bas** file
3. The JsonConverter module will appear in your project tree

### Step 4: Add Authentication Module

1. In VBA Editor, go to **Insert** → **Module**
2. Rename the module to **AuthModule** (optional)
3. Copy and paste the authentication code into the module

### Step 5: Create UserForms

#### UserForm1 (Login Form):
1. **Insert** → **UserForm**
2. Add controls:
   - `TextBox1` - Username input
   - `TextBox2` - Password input (set `PasswordChar` to `*`)
   - `CommandButton1` - Login button
3. Set form properties:
   - Name: `UserForm1`
   - Caption: `Autholas Login`

#### UserForm2 (Main Application):
1. **Insert** → **UserForm**
2. Design your main application interface
3. Set form properties:
   - Name: `UserForm2`
   - Caption: `Main Application`

### Step 6: Add References

1. In VBA Editor, go to **Tools** → **References**
2. Check the following references:
   - ✅ **Microsoft Scripting Runtime** (for Dictionary)
   - ✅ **Microsoft XML, v6.0** (for XMLHTTP)

## Project Structure

```
ExcelAuthProject.xlsm
├── Modules
│   ├── AuthModule          # Authentication functions
│   └── JsonConverter       # JSON library (imported)
├── UserForms
│   ├── UserForm1          # Login interface
│   └── UserForm2          # Main application
└── ThisWorkbook           # Workbook events (optional)
```

## Configuration

### 1. Set Your API Key

Edit the `AuthModule` and replace the API key:

```vba
Const API_KEY As String = "your_actual_api_key_here"
```

⚠️ **Important**: Never share your API key publicly or commit it to version control.

### 2. Update API URL (if needed)

```vba
Const API_URL As String = "https://autholas.nicholasdevs.my.id/api/auth"
```

### 3. Customize Device Name (Optional)

In the `AuthenticateUser` function, modify:

```vba
deviceName = "Custom Device Name"
```

## Usage

### Running the Application

1. **From Excel:**
   - Press `Alt + F8`
   - Select a macro that shows UserForm1
   - Or add a button on worksheet: **Developer** → **Insert** → **Button**

2. **From VBA Editor:**
   - Press `F5` while UserForm1 is selected
   - Or run this code in Immediate Window:
     ```vba
     UserForm1.Show
     ```

3. **Auto-start on Workbook Open:**
   
   Add to `ThisWorkbook`:
   ```vba
   Private Sub Workbook_Open()
       UserForm1.Show
   End Sub
   ```

### Login Flow

```
1. User opens Excel workbook
2. UserForm1 (Login) appears
3. User enters username and password
4. Click Login button
5. System authenticates with Autholas API
6. On success: UserForm2 (Main App) opens
7. On failure: Error message displayed
```

### Example Login Session

```
┌─────────────────────────────────────┐
│      Autholas Login System          │
├─────────────────────────────────────┤
│ Username: [your_username]           │
│ Password: [••••••••]                │
│                                     │
│         [Login Button]              │
└─────────────────────────────────────┘

✓ Login successful!
  Welcome, your_username!
```

## Code Explanation

### Main Functions

#### `AuthenticateUser(username, password) As Boolean`
- Performs user authentication
- Generates hardware ID
- Sends POST request to Autholas API
- Handles API responses
- Returns True on success, False on failure

#### `HandleAuthError(errorCode, Optional errorMessage)`
- Displays user-friendly error messages
- Handles various authentication error codes
- Shows appropriate MsgBox with title and description

### Authentication Flow

```vba
1. Collect username and password
2. Generate Hardware ID (HWID) from computer name
3. Build JSON payload with credentials
4. Send POST request to Autholas API
5. Parse JSON response
6. Handle authentication result:
   - Success (200): Show UserForm2
   - Unauthorized (401): Display error
   - Forbidden (403): Device/user banned
   - Other: Server error
```

## Error Handling

The system handles various authentication scenarios:

### Authentication Errors

| Error Code | Title | Description |
|------------|-------|-------------|
| `INVALID_CREDENTIALS` | Login Failed | Wrong username/password |
| `USER_BANNED` | Account Banned | User account suspended |
| `SUBSCRIPTION_EXPIRED` | Subscription Expired | Subscription period ended |
| `MAX_DEVICES_REACHED` | Device Limit Reached | Too many devices registered |
| `HWID_BANNED` | Device Banned | Hardware ID is blacklisted |
| `INVALID_API_KEY` | Service Error | Invalid API configuration |
| `RATE_LIMIT_EXCEEDED` | Too Many Attempts | Request rate limit hit |
| `DEVELOPER_SUSPENDED` | Service Unavailable | API developer suspended |
| `MISSING_PARAMETERS` | Invalid Request | Required fields missing |

### Network Errors

```vba
On Error GoTo ErrHandler
    ' ... authentication code ...
ErrHandler:
    MsgBox "Error saat request: " & Err.Description, vbCritical
```

## Troubleshooting

### Common Issues and Solutions

#### 1. "Compile Error: Can't find project or library"

**Solution:**
- Go to **Tools** → **References**
- Uncheck missing references (marked as "MISSING:")
- Check required references listed above
- Click OK and try again

#### 2. "Run-time error '429': ActiveX component can't create object"

**Solution:**
- Ensure XMLHTTP reference is added
- Try using `MSXML2.ServerXMLHTTP` instead:
  ```vba
  Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  ```

#### 3. JSON Parsing Errors

**Solution:**
- Verify JsonConverter.bas is imported correctly
- Check if response is valid JSON in Debug.Print
- Test with: `Debug.Print responseText`

#### 4. "Object doesn't support this property or method"

**Solution:**
- Ensure Scripting Runtime is referenced
- Check Dictionary syntax:
  ```vba
  Set payload = CreateObject("Scripting.Dictionary")
  ```

#### 5. Connection Timeout

**Solution:**
- Check internet connection
- Verify API URL is accessible
- Add timeout handling:
  ```vba
  http.setTimeouts 5000, 5000, 10000, 10000
  ```

#### 6. HWID Always Same/Generic

**Solution:**
- `COMPUTERNAME` environment variable might be empty
- Alternative HWID generation:
  ```vba
  hwid = Environ("USERNAME") & "_" & Environ("COMPUTERNAME")
  ```

## Security Features

- **Hardware ID Generation**: Uses computer name as unique identifier
- **Secure Communication**: HTTPS-only API calls
- **Password Masking**: TextBox `PasswordChar` property set to `*`
- **Error Sanitization**: Sensitive info not displayed in error messages
- **API Key Protection**: Stored as constant (consider encryption for production)

## Best Practices

### 1. Secure API Key Storage

For production, consider encrypting the API key:

```vba
' Store encrypted key
Const ENCRYPTED_KEY As String = "encrypted_value_here"

Function GetAPIKey() As String
    GetAPIKey = DecryptString(ENCRYPTED_KEY)
End Function
```

### 2. Enhanced HWID Generation

```vba
Function GetAdvancedHWID() As String
    Dim hwid As String
    hwid = Environ("COMPUTERNAME") & "_" & _
           Environ("USERNAME") & "_" & _
           Environ("PROCESSOR_IDENTIFIER")
    GetAdvancedHWID = hwid
End Function
```

### 3. Loading Indicators

```vba
Private Sub CommandButton1_Click()
    ' Show loading
    Me.CommandButton1.Enabled = False
    Me.CommandButton1.Caption = "Authenticating..."
    
    success = AuthenticateUser(user, pass)
    
    ' Reset button
    Me.CommandButton1.Enabled = True
    Me.CommandButton1.Caption = "Login"
End Sub
```

### 4. Session Management

Store session data globally:

```vba
' In a standard module
Public SessionToken As String
Public SessionExpiry As Date
Public AuthenticatedUser As String
```

## Advanced Configuration

### Custom HTTP Headers

```vba
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "User-Agent", "Excel-VBA-Client/1.0"
http.setRequestHeader "Accept", "application/json"
```

### Retry Logic

```vba
Function AuthenticateUserWithRetry(username As String, password As String, Optional maxRetries As Integer = 3) As Boolean
    Dim attempt As Integer
    For attempt = 1 To maxRetries
        If AuthenticateUser(username, password) Then
            AuthenticateUserWithRetry = True
            Exit Function
        End If
        Application.Wait Now + TimeValue("00:00:02")
    Next attempt
    AuthenticateUserWithRetry = False
End Function
```

### Logging

```vba
Sub LogAuthentication(username As String, success As Boolean)
    Dim logSheet As Worksheet
    Set logSheet = ThisWorkbook.Sheets("AuthLog")
    
    With logSheet
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        .Cells(lastRow, 1).Value = Now
        .Cells(lastRow, 2).Value = username
        .Cells(lastRow, 3).Value = IIf(success, "Success", "Failed")
        .Cells(lastRow, 4).Value = Environ("COMPUTERNAME")
    End With
End Sub
```

## Testing

### Manual Testing Checklist

- ✅ Valid credentials → Successful login
- ✅ Invalid credentials → Error message displayed
- ✅ Empty fields → Validation error
- ✅ Network disconnected → Connection error
- ✅ Wrong API key → Service error
- ✅ Banned account → Account banned message
- ✅ Expired subscription → Subscription error

### Debug Mode

Enable detailed logging:

```vba
' Add at the beginning of AuthenticateUser
Debug.Print "=== Authentication Started ==="
Debug.Print "Username: " & username
Debug.Print "HWID: " & hwid
Debug.Print "Request: " & jsonRequest
Debug.Print "Response: " & responseText
Debug.Print "Status: " & http.Status
Debug.Print "=== Authentication Ended ==="
```

View output in **Immediate Window** (`Ctrl + G`)

## Dependencies

- **VBA-JSON** (VBA-Tools): JSON parsing and serialization
  - GitHub: https://github.com/VBA-tools/VBA-JSON
  - License: MIT
- **Microsoft Scripting Runtime**: Dictionary object
- **Microsoft XML v6.0**: HTTP client functionality

## Compatibility

- **Excel Versions**: 2010, 2013, 2016, 2019, 365
- **Operating Systems**: Windows 7, 8, 10, 11
- **Architecture**: 32-bit and 64-bit Excel
- **VBA Version**: VBA 7.0+

## Deployment

### Distribution Steps

1. **Prepare workbook:**
   - Remove test data
   - Clear debug statements
   - Protect VBA code (Tools → VBAProject Properties → Protection)

2. **Code signing (optional):**
   - Use digital certificate
   - Sign VBA project for trust

3. **Package files:**
   - Include README.md
   - Add installation guide
   - Provide support contact

4. **Distribution:**
   - Share .xlsm file
   - Users enable macros when opening
   - Users configure API key

### Macro Security Settings

Users need to enable macros:
1. **File** → **Options** → **Trust Center**
2. **Trust Center Settings** → **Macro Settings**
3. Select: "Enable all macros" (for testing) or "Disable all macros with notification"

## Known Limitations

- Requires macro-enabled Excel (.xlsm)
- Needs internet connection for authentication
- HWID based on computer name (can be spoofed)
- No built-in token refresh mechanism
- API key stored in plain text in code

## Roadmap

- [ ] Encrypted API key storage
- [ ] Token refresh functionality
- [ ] Remember me feature
- [ ] Multi-language support
- [ ] Enhanced HWID generation
- [ ] Offline mode with cached credentials
- [ ] Two-factor authentication support

## Contributing

1. Fork the project
2. Create feature branch (`git checkout -b feature/AmazingFeature`)
3. Follow VBA naming conventions
4. Add comments for complex logic
5. Commit changes (`git commit -m 'Add AmazingFeature'`)
6. Push to branch (`git push origin feature/AmazingFeature`)
7. Open Pull Request

## License

This project is provided as-is for educational and development purposes. Please respect the terms of service of the Autholas API.

## Support

For issues related to:
- **Autholas API**: Contact Autholas support at support@nicholasdevs.my.id
- **VBA-JSON Library**: https://github.com/VBA-tools/VBA-JSON/issues
- **Excel VBA**: Microsoft Office support documentation
- **This implementation**: Open an issue in this repository

## Acknowledgments

- **VBA-Tools** for the excellent VBA-JSON library
- **Autholas** for the authentication API service
- **Microsoft** for Excel VBA platform

## Changelog

### v1.0.0 (Initial Release)
- ✅ Basic authentication implementation
- ✅ Hardware ID generation
- ✅ Error handling system
- ✅ UserForm integration
- ✅ JSON serialization/parsing
- ✅ Multi-device support

---
