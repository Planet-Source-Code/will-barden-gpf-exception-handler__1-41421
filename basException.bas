Attribute VB_Name = "basException"
Option Explicit

' -------------------------------------------------------------- '
' module to handle unhandled exceptions (GPFs)
' created 25/11/02
' modified  10/12/02
' will barden
'
' 10/12/02 - added a more descriptive error message, and
'            setup the error handler to use VBs's internal
'            error bubbling to raise it.
' -------------------------------------------------------------- '

' -------------------------------------------------------------- '
' apis
' -------------------------------------------------------------- '

' used to set and remove our callback
Public Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long

' to raise a GPF (for testing)
Public Declare Sub RaiseException Lib "kernel32" (ByVal dwExceptionCode As Long, ByVal dwExceptionFlags As Long, ByVal nNumberOfArguments As Long, lpArguments As Long)

' to get the last GPF code
Public Declare Function GetExceptionInformation Lib "kernel32" () As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

' -------------------------------------------------------------- '
' consts
' -------------------------------------------------------------- '

' return values from our callback
Public Const EXCEPTION_CONTINUE_EXECUTION = -1
Public Const EXCEPTION_CONTINUE_SEARCH = 0

' length field in the EXCEPTION_RECORD struct
Public Const EXCEPTION_MAXIMUM_PARAMETERS = 15

' to describe the violation - defined in windows.h
Public Const EXCEPTION_CONTINUABLE              As Long = &H0
Public Const EXCEPTION_NONCONTINUABLE           As Long = &H1

Public Const EXCEPTION_ACCESS_VIOLATION         As Long = &HC0000005 ' The thread tried to read from or write to a virtual address for which it does not have the appropriate access
Public Const EXCEPTION_BREAKPOINT               As Long = &H80000003 ' A breakpoint was encountered.
Public Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED    As Long = &HC000008C ' The thread tried to access an array element that is out of bounds and the underlying hardware supports bounds checking.
Public Const EXCEPTION_FLT_DIVIDE_BY_ZERO       As Long = &HC000008E ' The thread tried to divide a floating-point value by a floating-point divisor of zero.
Public Const EXCEPTION_FLT_INVALID_OPERATION    As Long = &HC0000090 ' This exception represents any floating-point exception not included in this list
Public Const EXCEPTION_FLT_OVERFLOW             As Long = &HC0000091 ' The exponent of a floating-point operation is greater than the magnitude allowed by the corresponding type
Public Const EXCEPTION_INT_DIVIDE_BY_ZERO       As Long = &HC0000094 ' The thread tried to divide an integer value by an integer divisor of zero.
Public Const EXCEPTION_INT_OVERFLOW             As Long = &HC0000095 ' The result of an integer operation caused a carry out of the most significant bit of the result
Public Const EXCEPTION_ILLEGAL_INSTRUCTION      As Long = &HC000001D ' The thread tried to execute an invalid instruction
Public Const EXCEPTION_PRIV_INSTRUCTION         As Long = &HC0000096 ' The thread tried to execute an instruction whose operation is not allowed in the current machine mode

' -------------------------------------------------------------- '
' structs
' -------------------------------------------------------------- '

' holds info about a specific eception
Public Type EXCEPTION_RECORD
  ExceptionCode      As Long  ' type of exception - defined above
  ExceptionFlags     As Long  ' whether the exception is continuable or not
  pExceptionRecord   As Long  ' pointer to another EXCEPTION_RECORD struct (for nested exceptions)
  ExceptionAddress   As Long  ' the address at which the exception occurred
  NumberParameters   As Long  ' number of params in the following array
  Information(EXCEPTION_MAXIMUM_PARAMETERS - 1) As Long ' extra info.. not really needed.
End Type

' processor specific - not really needed anyway
Public Type CONTEXT
  Null               As Long
End Type

' wrapper for the above types
Public Type EXCEPTION_POINTERS
  pExceptionRecord   As EXCEPTION_RECORD
  ContextRecord      As CONTEXT
End Type

' -------------------------------------------------------------- '
' private variables
' -------------------------------------------------------------- '
Private mlpOldProc As Long

' -------------------------------------------------------------- '
' methods
' -------------------------------------------------------------- '

' setup the new handler
Public Function StartGPFHandler() As Boolean
   
   ' assume success
   StartGPFHandler = True
   
   ' if we're already handling, there's no point
   If mlpOldProc = 0 Then
   
      ' set up the handler
      mlpOldProc = SetUnhandledExceptionFilter(AddressOf ExceptionHandler)
      ' not all systems will return a handle
      If mlpOldProc = 0 Then mlpOldProc = 1
      
   End If
   
End Function

' release the new handler
Public Sub StopGPFHandler()
   
   ' release the handler
   SetUnhandledExceptionFilter vbNull
   
   ' reset the variable
   mlpOldProc = 0
   
End Sub

' just for debugging - test the handler by firing a GPF
Public Sub TestGPFHandler()

   ' raise a GPF
   RaiseException EXCEPTION_ARRAY_BOUNDS_EXCEEDED, 0, 0, 0
   
End Sub

' altered on 10/12/02 by request - this function now simply raises
' an error so that VB can handle it properly, via On Error.
Public Function ExceptionHandler(ByRef uException As EXCEPTION_POINTERS) As Long
Dim lTmp       As Long
Dim sType      As String
Dim lAddress   As Long
Dim sContinue  As String

   ' let's get some information about the error in order
   ' to raise a nicely defined, and explanatory error via VB
   CopyMemory lTmp, ByVal uException.pExceptionRecord.ExceptionCode, 4
   Select Case lTmp
      Case EXCEPTION_ACCESS_VIOLATION
         sType = "EXCEPTION_ACCESS_VIOLATION"
      Case EXCEPTION_BREAKPOINT
         sType = "EXCEPTION_BREAKPOINT"
      Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
         sType = "EXCEPTION_ARRAY_BOUNDS_EXCEEDED"
      Case EXCEPTION_FLT_DIVIDE_BY_ZERO
         sType = "EXCEPTION_FLT_DIVIDE_BY_ZERO"
      Case EXCEPTION_FLT_INVALID_OPERATION
         sType = "EXCEPTION_FLT_INVALID_OPERATION"
      Case EXCEPTION_FLT_OVERFLOW
         sType = "EXCEPTION_FLT_OVERFLOW"
      Case EXCEPTION_INT_DIVIDE_BY_ZERO
         sType = "EXCEPTION_INT_DIVIDE_BY_ZERO"
      Case EXCEPTION_INT_OVERFLOW
         sType = "EXCEPTION_INT_OVERFLOW"
      Case EXCEPTION_ILLEGAL_INSTRUCTION
         sType = "EXCEPTION_ILLEGAL_INSTRUCTION"
      Case EXCEPTION_PRIV_INSTRUCTION
         sType = "EXCEPTION_PRIV_INSTRUCTION"
      Case Else
         sType = "Unknown exception type"
   End Select

   ' check for a couple of other important points..
   With uException.pExceptionRecord
      ' can we continue from this error?
      If .ExceptionFlags = EXCEPTION_CONTINUABLE Then
         sContinue = "Ok to continue."
      ElseIf .ExceptionFlags = EXCEPTION_NONCONTINUABLE Then
         sContinue = "NOT ok to continue."
      Else
         sContinue = "Probably safe to continue, but better not."
      End If
      ' and lastly, where the error occurred.
      lAddress = .ExceptionAddress
   End With

   ' raise the error so that our user can handle it via VB
   Err.Raise vbObjectError + 513, _
             "Exception Handler", _
             "An unhandled error (" & sType & ") " & vbCrLf & _
               "occurred at: " & lAddress & ". " & sContinue

   ' continue with execution
   ExceptionHandler = EXCEPTION_CONTINUE_EXECUTION
   
End Function

