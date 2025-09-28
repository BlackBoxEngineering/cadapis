'////////////////////////////////////////////////////////////////////////////////////////
'// EDUCATIONAL SECURITY DEMONSTRATION
'// Privilege Escalation and System Shutdown from CAD API Context
'//
'// Author: Matt McPhillips
'// Email: mattmcp@blackboxengineering.co.uk
'//
'// This code demonstrates how a seemingly innocent CAD macro can escalate privileges
'// and perform system-level operations. Originally developed as a security research
'// demonstration, this technique shows that VBA macros can access Windows APIs
'// for privilege manipulation and system control.
'//
'// WARNING: This is for security research and educational purposes only.
'// Do not use in production systems. BlackBox Engineering - Security Research.
'//
'////////////////////////////////////////////////////////////////////////////////////////

Option Explicit

'////////////////////////////////////////////////////////////////////////////////////////
'// Windows API Structures and Constants
'//
' LUID (Locally Unique Identifier) - issued with every Windows boot
Public Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type

' TOKEN_PRIVILEGES - Structure for retrieving user privilege information
Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheUserLogonIdentifier As LUID
    Attributes As Long
End Type

' Shutdown command constants
Public Const EWX_SHUTDOWN As Long = 1
Public Const EWX_FORCE As Long = 4
Public Const EWX_REBOOT = 2

'////////////////////////////////////////////////////////////////////////////////////////
'// Windows API Function Declarations
'//
'// These functions access Windows system libraries and demonstrate how VBA can call
'// low-level system functions to perform privileged operations
'//

'// 64-bit compatibility: Use conditional compilation for VBA7 (Office 2010+)
#If VBA7 Then
    '// 64-bit compatible declarations using PtrSafe and LongPtr
    Public Declare PtrSafe Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
    Public Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
    Public Declare PtrSafe Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As LongPtr, ByVal DesiredAccess As Long, TokenHandle As LongPtr) As Long
    Public Declare PtrSafe Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
    Public Declare PtrSafe Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As LongPtr, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
#Else
    '// Legacy 32-bit declarations for older Office versions
    Public Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
    Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
    Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
    Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
    Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
#End If

'////////////////////////////////////////////////////////////////////////////////////////
'// SECURITY ANALYSIS OF API CHAIN:
'//
'// 1. GetCurrentProcess() - Gets handle to our process
'// 2. OpenProcessToken() - Opens our process's security token for modification
'// 3. LookupPrivilegeValue() - Finds the shutdown privilege identifier
'// 4. AdjustTokenPrivileges() - GRANTS shutdown privilege to our token
'// 5. ExitWindowsEx() - Uses newly acquired privilege to shutdown system
'//
'// This demonstrates a complete privilege escalation attack chain using
'// legitimate Windows APIs in an unintended way from within a CAD macro.
'////////////////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////////////////////////////////
'// Main Demonstration Function
'//
Public Sub DemonstratePrivilegeEscalation()
    MsgBox "This demonstrates privilege escalation from a CAD macro context." & vbCrLf & _
           "The macro will attempt to gain shutdown privileges and offer system control.", _
           vbInformation, "Security Demonstration"
    
    ' Attempt privilege escalation
    If AcquireShutdownPrivileges Then
        MsgBox "SUCCESS: Privileges escalated successfully!" & vbCrLf & _
               "The macro now has system shutdown capabilities.", _
               vbExclamation, "Privilege Escalation Successful"
        
        ' Offer the "classic" choice (but safer for demo)
        OfferShutdownDemo
    Else
        MsgBox "FAILED: Could not escalate privileges." & vbCrLf & _
               "This may be due to system security policies.", _
               vbCritical, "Privilege Escalation Failed"
    End If
End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'// Privilege Escalation Function
'//
Private Function AcquireShutdownPrivileges() As Boolean
    On Error GoTo ErrorHandler
    
    ' Token access constants
    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8
    Const SE_PRIVILEGE_ENABLED = &H2
    
    ' Variables for privilege manipulation (64-bit compatible)
    #If VBA7 Then
        Dim hdlProcessHandle As LongPtr
        Dim hdlTokenHandle As LongPtr
    #Else
        Dim hdlProcessHandle As Long
        Dim hdlTokenHandle As Long
    #End If
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    
    ' Step 1: Get handle to current process
    hdlProcessHandle = GetCurrentProcess
    
    ' Step 2: Open the process token with required access rights
    If OpenProcessToken(hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle) = 0 Then
        AcquireShutdownPrivileges = False
        Exit Function
    End If
    
    ' Step 3: Look up the shutdown privilege value
    If LookupPrivilegeValue("", "SeShutdownPrivilege", tmpLuid) = 0 Then
        AcquireShutdownPrivileges = False
        Exit Function
    End If
    
    ' Step 4: Set up the privilege structure
    tkp.PrivilegeCount = 1
    tkp.TheUserLogonIdentifier = tmpLuid
    tkp.Attributes = SE_PRIVILEGE_ENABLED
    
    ' Step 5: Adjust the token privileges
    If AdjustTokenPrivileges(hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded) <> 0 Then
        AcquireShutdownPrivileges = True
    Else
        AcquireShutdownPrivileges = False
    End If
    
    Exit Function
    
ErrorHandler:
    AcquireShutdownPrivileges = False
End Function

'////////////////////////////////////////////////////////////////////////////////////////
'// The "Classic" Shutdown Demo with Progressive Alerts
'//
Private Sub OfferShutdownDemo()
    Dim response As VbMsgBoxResult
    
    ' Show the progression of the attack
    MsgBox "The macro has taken your token", vbCritical, "Security Alert 1/3"
    MsgBox "The macro owns your computer", vbCritical, "Security Alert 2/3"
    MsgBox "The macro is trying to shut your computer down", vbCritical, "Security Alert 3/3"
    
    ' Give them the final choice
    response = MsgBox("Continue with shut down or exit?", vbYesNo + vbCritical, "Final Warning")
    
    If response = vbYes Then
        ' Execute the shutdown sequence
        ExecuteShutdown
    Else
        MsgBox "Shutdown cancelled. The macro demonstrated privilege escalation but chose not to execute." & vbCrLf & vbCrLf & _
               "This shows how a malicious macro could take control of your system.", _
               vbInformation, "Educational Demo Complete"
    End If
End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'// Execute System Shutdown (Original Implementation)
'//
Private Sub ExecuteShutdown()
    ' Display final warning
    MsgBox "System shutdown initiated..." & vbCrLf & _
           "This will force shutdown and reboot your machine.", _
           vbCritical, "Shutdown Warning"
    
    ' Execute the shutdown command with force and reboot
    ExitWindowsEx (EWX_SHUTDOWN Or EWX_FORCE Or EWX_REBOOT), &HFFFF
End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'// Main Entry Point (Original Functionality)
'//
Sub Main()
    MsgBox "WARNING: This is a privilege escalation and shutdown demonstration!" & vbCrLf & vbCrLf & _
           "This macro will attempt to gain system privileges and" & vbCrLf & _
           "may shut down your computer if you continue." & vbCrLf & vbCrLf & _
           "BlackBox Engineering - Security Research", _
           vbExclamation, "Security Demonstration - LIVE"
    
    DemonstratePrivilegeEscalation
End Sub

'////////////////////////////////////////////////////////////////////////////////////////
'// EDUCATIONAL NOTES:
'//
'// This code demonstrates several important security concepts:
'//
'// 1. PRIVILEGE ESCALATION: How applications can request elevated privileges
'// 2. TOKEN MANIPULATION: Working with Windows security tokens
'// 3. API ABUSE: Using legitimate APIs for unintended purposes
'// 4. SOCIAL ENGINEERING: User manipulation techniques
'// 5. UNEXPECTED ATTACK VECTORS: Security threats from CAD macros
'//
'// Modern mitigations include:
'// - User Account Control (UAC)
'// - Code signing requirements
'// - Macro security settings
'// - Application sandboxing
'// - Privilege separation
'//
'// This serves as a reminder that security threats can come from
'// unexpected sources and that proper input validation, privilege
'// management, and user education are essential.
'//
'// WARNING: This macro WILL shut down your computer if you select 'No'
'// when prompted. Use only in controlled environments for education.
'//
'////////////////////////////////////////////////////////////////////////////////////////