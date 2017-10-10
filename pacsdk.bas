Attribute VB_Name = "pacsdk"
'   System Functions
Declare Function pac_GetModuleType Lib "PACSDK_vb.dll" (ByVal slot As Byte) As Integer
Declare Function pac_GetModuleName Lib "PACSDK_vb.dll" (ByVal slot As Long, ByVal strName As String) As Integer
Declare Function pac_GetRotaryID Lib "PACSDK_vb.dll" () As Integer
Declare Sub pac_GetSerialNumber Lib "PACSDK_vb.dll" (ByVal SerialNumber As String)
Declare Sub pac_GetSDKVersion Lib "PACSDK_vb.dll" (ByVal sdk_version As String)
Declare Sub pac_ChangeSlot Lib "PACSDK_vb.dll" (ByVal slotNo As Byte)
Declare Function pac_CheckSDKVersion Lib "PACSDK_vb.dll" (ByVal version As Long) As Boolean

'Backplane API
Declare Function pac_GetDIPSwitch Lib "PACSDK_vb.dll" () As Integer
Declare Function pac_GetSlotCount Lib "PACSDK_vb.dll" () As Integer
Declare Sub pac_GetBackplaneID Lib "PACSDK_vb.dll" (ByVal backplane_version As String)
Declare Function pac_GetBatteryLevel Lib "PACSDK_vb.dll" (ByVal nBattery As Long) As Integer
Declare Function pac_ModuleExists Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long) As Boolean
Declare Function pac_EnableRetrigger Lib "PACSDK_vb.dll" (ByVal iValue As Byte)


'Memory API
Declare Function pac_GetMemorySize Lib "PACSDK_vb.dll" (ByVal mem_type As Long) As Long
Declare Function pac_ReadMemory Lib "PACSDK_vb.dll" (ByVal address As Long, ByRef lpBuffer As Byte, ByVal dwLength As Long, ByVal mem_type As Long) As Boolean
Declare Function pac_WriteMemory Lib "PACSDK_vb.dll" (ByVal address As Long, ByRef lpBuffer As Byte, ByVal dwLength As Long, ByVal mem_type As Long) As Boolean
Declare Sub pac_EnableEEPROM Lib "PACSDK_vb.dll" (ByVal bEnable As Boolean)


'Watch Dog Timer Functions
Declare Function pac_EnableWatchDog Lib "PACSDK_vb.dll" (ByVal wdt As Long, ByVal value As Long) As Boolean
Declare Sub pac_DisableWatchDog Lib "PACSDK_vb.dll" (ByVal wdt As Long)
Declare Sub pac_RefreshWatchDog Lib "PACSDK_vb.dll" (ByVal wdt As Long)
Declare Function pac_GetWatchDogState Lib "PACSDK_vb.dll" (ByVal wdt As Long) As Boolean
Declare Function pac_GetWatchDogTime Lib "PACSDK_vb.dll" (ByVal wdt As Long) As Long
Declare Function pac_SetWatchDogTime Lib "PACSDK_vb.dll" (ByVal wdt As Long, ByVal value As Long) As Boolean


'UART API
Declare Function uart_Open Lib "PACSDK_vb.dll" (ByVal ConnectionString As String) As Long
Declare Function uart_Close Lib "PACSDK_vb.dll" (ByVal hPort As Long) As Boolean
Declare Function uart_Send Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal buf As String) As Boolean
Declare Function uart_Recv Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal buf As String) As Boolean
Declare Function uart_SendCmd Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal cmd As String, ByVal szResult As String) As Boolean
Declare Function uart_SetTimeOut Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal msec As Long, ByVal ctoType As Long) As Boolean
Declare Function uart_EnableCheckSum Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal bEnable As Boolean) As Boolean
Declare Function uart_SetTerminator Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal szTerm As String) As Boolean


'Slot Interrupt API


Declare Function pac_UnregisterSlotInterrupt Lib "PACSDK_vb.dll" (ByVal slot As Byte) As Boolean
Declare Function pac_SetSlotInterruptPriority Lib "PACSDK_vb.dll" (ByVal slot As Byte, ByVal nPriority As Long) As Boolean
Declare Function pac_EnableSlotInterrupt Lib "PACSDK_vb.dll" (ByVal slot As Byte, ByVal bEnable As Boolean)
Declare Function pac_InterruptDone Lib "PACSDK_vb.dll" (ByVal slot As Byte)
Declare Function pac_GetSlotInterruptEvent Lib "PACSDK_vb.dll" (ByVal slot As Byte)
Declare Function pac_GetSlotInterruptID Lib "PACSDK_vb.dll" (ByVal slot As Byte) As Long
Declare Function pac_SetTriggerType Lib "PACSDK_vb.dll" (ByVal iType As Long)
Declare Function pac_InterruptInitialize Lib "PACSDK_vb.dll" (ByVal slot As Byte) As Boolean
Declare Function pac_SetSlotInterruptEvent Lib "PACSDK_vb.dll" (ByVal slot As Byte, ByVal hEvent As Long)


'Error Handling API
Declare Function pac_GetLastError Lib "PACSDK_vb.dll" () As Long
Declare Sub pac_SetLastError Lib "PACSDK_vb.dll" (ByVal errorno As Long)
Declare Sub pac_ClearLastError Lib "PACSDK_vb.dll" ()
Declare Sub pac_GetErrorMessage Lib "PACSDK_vb.dll" (ByVal dwMessageID As Long, ByVal lpBuffer As String)


'PAC Local/Remote IO API
Declare Function pac_WriteDO Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iDO_TotalCh As Long, ByVal lDO_Value As Long) As Boolean
Declare Function pac_WriteDOBit Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iDO_TotalCh As Long, ByVal iChannel As Long, ByVal iBitValue As Long) As Boolean
Declare Function pac_ReadDO Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iDO_TotalCh As Long, ByRef lDO_Value As Long) As Boolean
Declare Function pac_ReadDI Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iDI_TotalCh As Long, ByRef lDI_Value As Long) As Boolean
Declare Function pac_ReadDIO Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iDI_TotalCh As Long, ByVal iDO_TotalCh As Long, ByRef lDI_Value As Long, ByRef lDO_Value As Long) As Boolean
Declare Function pac_ReadDILatch Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iDI_TotalCh As Long, ByVal iLatchType As Long, ByRef lDI_Latch_Value As Long) As Boolean
Declare Function pac_ClearDILatch Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long) As Boolean
Declare Function pac_ReadDIOLatch Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iDI_TotalCh As Long, ByVal iDO_TotalCh As Long, ByVal iLatchType As Long, ByRef lDI_Latch_Value As Long, ByRef lDO_Latch_Value As Long) As Boolean
Declare Function pac_ClearDIOLatch Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long) As Boolean
Declare Function pac_ReadDICNT Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iChannel As Long, ByVal iDI_TotalCh As Long, ByRef lCounter_Value As Long) As Boolean
Declare Function pac_ClearDICNT Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iChannel As Long, ByVal iDI_TotalCh As Long) As Boolean
Declare Function pac_WriteAO Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iChannel As Long, ByVal iAO_TotalCh As Long, ByVal fValue As Single) As Boolean
Declare Function pac_ReadAO Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iChannel As Long, ByVal iAO_TotalCh As Long, ByRef fValue As Single) As Boolean
Declare Function pac_ReadAI Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iChannel As Long, ByVal iAI_TotalCh As Long, ByRef fValue As Single) As Boolean
Declare Function pac_ReadAIHex Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iChannel As Long, ByVal iAI_TotalCh As Long, ByRef iValue As Long) As Boolean
Declare Function pac_ReadAIAll Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByRef fValue As Single) As Boolean
Declare Function pac_ReadAIAllHex Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByRef iValue As Long) As Boolean
Declare Function pac_ReadCNT Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iChannel As Long, ByRef lCounter_Value As Long) As Boolean
Declare Function pac_ClearCNT Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iChannel As Long) As Boolean
Declare Function pac_ReadCNTOverflow Lib "PACSDK_vb.dll" (ByVal hPort As Long, ByVal slot As Long, ByVal iChannel As Long, ByRef iOverflow As Long) As Boolean


'#define PAC_REMOTE_IO_BASE      (0x1000)
'#define PAC_REMOTE_IO(iAddress) (PAC_REMOTE_IO_BASE+iAddress)
Public Function PAC_REMOTE_IO(iAddress As Integer) As Integer
    PAC_REMOTE_IO = CInt("&H" + "1000") + iAddress
End Function


'#define pac_GetBit(v, ndx) (v & (1<<ndx))
Public Function pac_GetBit(value As Integer, index As Integer) As Boolean
    Dim bit As Integer
    bit = (value And (10 ^ (index - 1)))
    
    If bit = 0 Then pac_GetBit = False
    
    If bit = 1 Then pac_GetBit = True

End Function



