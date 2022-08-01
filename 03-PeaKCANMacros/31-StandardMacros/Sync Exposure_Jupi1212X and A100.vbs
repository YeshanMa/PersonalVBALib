'------------------------------------------------------------------------------
'Program:	FILE DESCRIPTION: Sync Exposure of Jupi1212X and A10
'Author: 	Ma,Yeshan
'Date: 		2022-07-28
'Ver:		0.1

'Change History
'Ver 0.10	2022-07-26	1st Draft
'Ver 0.11	2022-07-28	TBD

'----------------------------------------------------------------
Option Explicit

'Define Variables

'------Boolean Variables for IOs to sync A100 and Jupi 1212X---------
Dim bOUT_FD_SYNC_IN As Boolean = FALSE
Dim bOUT_FD_XRAY_ON As Boolean = FALSE

Dim bOUT_HVG_SWR As Boolean = FALSE
Dim bOUT_HVG_SUMREL As Boolean = FALSE

Dim bIN_FD_EXPO_EN As Boolean = FALSE

Dim bIN_HS_HK As Boolean = FALSE
Dim bIN_HS_VK As Boolean = FALSE

'----Assign Variable to Signals over CAN Messages---
Set bOUT_FD_SYNC_IN   = Signals("bOUT_FD_SYNC_IN")
Set bOUT_FD_XRAY_ON   = Signals("bOUT_FD_XRAY_ON")
Set bOUT_HVG_SWR   = Signals("bOUT_HVG_SWR")
Set bOUT_HVG_SUMREL   = Signals("bOUT_HVG_SUMREL")
Set bIN_HS_HK   = Signals("bIN_HS_HK")
Set bIN_HS_VK   = Signals("bIN_HS_VK")
Set bIN_FD_EXPO_EN   = Signals("bIN_FD_EXPO_EN")


'------Variables to Define X-Ray Exposure Parameters-----------
Const nXray_Exposure_FrameRate As Integer = 20			'FrameRate set at 20 FPS
Const nXray_Exposure_PulseWidth As Integer = 10			'X-Ray exposure pulse width set at 10ms, Duty Cycle 10/(1000/20)= 20%

Const dHVG_SUMREL_SWR_DelayON_Timer As double = 500		'SUMREL ON and Delay for 500ms, then SWR is allowed to be ON.
Const dHVG_SWR_ON_Timer As double = nXray_Exposure_PulseWidth					'SWR On time for 10ms, which equals to X-Ray Exposure Pulse Width


'----CAN Messages Configration---
Dim bOUT_FD_SYNC_IN As Boolean = FALSE
Dim bOUT_FD_XRAY_ON As Boolean = FALSE


'-----Program Start--------
Sub XRaySync_HVG_FD()

'DESCRIPTION: Rest Bus msg counter
Dim UseConn, conn
    ' Find the first enabled connection in the project that uses the CAN protocol
Set UseConn = Nothing
    For Each conn In Connections
        If conn.IsEnabled And conn.Protocol = peProtocolCAN Then
            Set UseConn = conn
        Exit For
        End If
    Next 
    If UseConn Is Nothing Then
        MsgBox "Project does not contain any enabled CAN connections"
        Exit Sub
    End If
    
Dim MyClient, PcanConn
Set MyClient = CreateObject("PCAN3.PCANClient")
MyClient.Device = UseConn.Device
MyClient.Name = "PCANLight_USB_16"
Set PcanConn = MyClient.Connections.Add(UseConn.CommunicationObject.NetName)


Dim counter1, counter2, counter3, counter4
Dim loopcounter
  Dim i,j,k,l
  ' Now create and initialize a new transmit message
Dim msg,GCAN_IO_IN,GCAN_IO_OUT,StartGCANIO
Set msg= MyClient.Messages.Add
Set GCAN_IO_IN = MyClient.Messages.Add
Set GCAN_IO_OUT= MyClient.Messages.Add
Set StartGCANIO= MyClient.Messages.Add

timestamp = MyClient.GetSystemTime
  
With StartGCANIO
    .ID = &000
    .DLC = 2
    .Data(0) = &H01
    .Data(1) = &H03
End With

With GCAN_IO_IN
    .ID = &183
    .DLC = 1
    .Data(0) = &H00
End With

With GCAN_IO_OUT
    .ID = &203								
    .DLC = 1
    .Data(0) = &H00
End With

'----Initial Counter------
 i=0
 j=0
 k=0
 l=0

' Create a new client and connect it to the same Net that the
'  found connection uses

Dim MyClient, PcanConn
Set MyClient = CreateObject("PCAN3.PCANClient")
MyClient.Device = UseConn.Device
MyClient.Name = "PCANLight_USB_16"
Set PcanConn = MyClient.Connections.Add(UseConn.CommunicationObject.NetName)

'------Loop Start-------
 do
 	loopcounter=loopcounter+1
 	' // 20ms 
 	i=i+1
  	msg.Data(5) = i							'//Initializing the byte 5 with message counter for message1							
  	If i>15 then i=0 End If
    msg.Data(6) = GCAN_IO_IN.Data(0)+GCAN_IO_IN.Data(1)+GCAN_IO_IN.Data(2)+GCAN_IO_IN.Data(3)+ GCAN_IO_IN.Data(4)+GCAN_IO_IN.Data(5) '//Implementing a simple CheckSum
    msg.Write PcanConn, timestamp				'//Transmitting data into CAN bus at Cycletime= 20 ms
    
    counter1=counter1+1
    If(counter1=2)Then
    	j=j+1
        GCAN_IO_IN.Data(1) = j							'//Initializing the byte 1 with message counter for message1						
        if j>255 then j=0 End If
    	GCAN_IO_IN.Write PcanConn, timestamp				'//Transmitting data into CAN bus at Cycletime= 40 ms
    	counter1=0
    End If
    
    counter2=counter2+1
    If(counter2=4)Then
     	k=k+1
		GCAN_IO_OUT.Data(2) = k							'//Initializing the byte 2 with message counter for message1						     	
		if k>128 Then k=0 End If
    	GCAN_IO_OUT.Write PcanConn, timestamp				'//Transmitting data into CAN bus at Cycletime= 80 ms
    	counter2=0
    End If    
    
    counter3=counter3+1
    If(counter3=5)Then
		l=l+1
		StartGCANIO.Data(3) = l							'//Initializing the byte 3 with message counter for message1						     	
		if l>64 Then l=0 End If    
    	StartGCANIO.Write PcanConn, timestamp				'//Transmitting data into CAN bus at Cycletime= 100 ms
    	counter3=0
    End If    
    
    ' lowest timer intervall is 20ms - so we need a 20ms resolution
    Wait 20
    
 	
  	Loop While(loopcounter<1000) ' run forever?...no...until 1000 loops 
  
  
  ' Wait until all messages have been sent before finishing macro,
  '  since this would terminate the client and delete all messages
  '  that are still in the queue
  While not MyClient.XmtQueueEmpty
    Wait 500
  Wend
  Wait 500

  End Sub





'-------Below for Refer------

Sub NewClientSend()
'DESCRIPTION: Sends CAN messages using a new PCAN client
  Dim UseConn, conn
  ' Find the first enabled connection in the project that uses the CAN protocol
  Set UseConn = Nothing
  For Each conn In Connections
    If conn.IsEnabled And conn.Protocol = peProtocolCAN Then
      Set UseConn = conn
      Exit For
    End If
  Next 
  If UseConn Is Nothing Then
    MsgBox "Project does not contain any enabled CAN connections"
    Exit Sub
  End If

  ' Create a new client and connect it to the same Net that the
  '  found connection uses
  Dim MyClient, PcanConn
  Set MyClient = CreateObject("PCAN4.PCANClient")
  MyClient.Device = UseConn.Device
  MyClient.Name = "Macro"
  Set PcanConn = MyClient.Connections.Add(UseConn.CommunicationObject.NetName)

  ' Now create and initialize a new transmit message
  Dim msg, timestamp, i
  Set msg = MyClient.Messages.Add
  timestamp = MyClient.GetSystemTime + 1000
  With msg
    .ID = &H100
    .DLC = 4
    .MsgType = pcanMsgTypeExtended
    .Data(0) = &H11
    .Data(1) = &H22
    .Data(2) = &H33
  End With
  for i = 1 To 20
    msg.Data(3) = i
    msg.Write PcanConn, timestamp
    timestamp = timestamp + 500
  next
  ' Wait until all messages have been sent before finishing macro,
  '  since this would terminate the client and delete all messages
  '  that are still in the queue
  While not MyClient.XmtQueueEmpty
    Wait 500
  Wend
  Wait 500
End Sub


Sub WaitForID100()
'DESCRIPTION: Waits until CAN-ID 100h is received
  ' To view the output messages of this macro, open the Output Window and
  '  select the "Macro" tab
  Dim UseConn, conn
  ' Find the first enabled connection in the project that uses the CAN protocol
  Set UseConn = Nothing
  For Each conn In Connections
    If conn.IsEnabled And conn.Protocol = peProtocolCAN Then
      Set UseConn = conn
      Exit For
    End If
  Next 
  If UseConn Is Nothing Then
    MsgBox "Project does not contain any enabled CAN connections"
    Exit Sub
  End If

  ' Create a new client and connect it to the same Net that the
  '  found connection uses
  Dim MyClient, PcanConn
  Set MyClient = CreateObject("PCAN4.PCANClient")
  MyClient.Device = UseConn.Device
  MyClient.Name = "Macro"
  Set PcanConn = MyClient.Connections.Add(UseConn.CommunicationObject.NetName)
  If Not PcanConn.IsConnected Then
    MsgBox "Cannot connect to net " & PcanConn.NetName
    Exit Sub
  End If

  PcanConn.RegisterMsg &H100, &H200, False, False

  Dim RcvMsg, i
  i = 0
  Set RcvMsg = MyClient.Messages.Add
  Do
    Do While Not RcvMsg.Read
      ' Wait for an incoming message
      Wait 1
    Loop
    If RcvMsg.LastError = pcanErrorOk Then
      PrintToOutputWindow "Received!"
      i = i + 1
    End If
  Loop While RcvMsg.ID <> &H100
  PrintToOutputWindow "Finished, " & CStr(i) & " messages received!"
End Sub




  ' Create two tracers, which will be used alternately
  Set doc = Documents.Add(peDocumentKindTrace)
  Set wnd = doc.ActiveWindow
  Set tracer1 = wnd.Object.Tracer
  wnd.Left = 0
  wnd.Top = 0
  wnd.Height = 250
  wnd.Width = 600

  Set doc = Documents.Add(peDocumentKindTrace)
  Set wnd = doc.ActiveWindow
  Set tracer2 = wnd.Object.Tracer
  wnd.Left = 0
  wnd.Top = 250
  wnd.Height = 250
  wnd.Width = 600
  
  Set wnd = Nothing
  Set doc = Nothing

  ' Pre-configure both tracers
  tracer1.BufferType = peTraceBufferTypeFile
  tracer2.BufferType = peTraceBufferTypeFile

  ' Some initializations
  Set CurrentTracer = tracer1  ' Begin with tracer1
  Set NextTracer = tracer2

  TracerNumber = 1
  CurrentTracer.Document.Save DestDir & "\Trace" & TracerNumber & ".trc"
  CurrentTracer.Start
  IsRunning = True

  Do While IsRunning

    ' Tracer records data until the number of tracer entries reaches maximum
    Do While IsRunning And (CurrentTracer.EntryCount < MessagesPerTracer)
      Wait 50
      IsRunning = CurrentTracer.TraceState = peTraceStarted
    Loop
    
    If IsRunning Then
      TracerNumber = TracerNumber + 1
      NextTracer.Clear
      NextTracer.Document.Save DestDir & "\Trace" & TracerNumber & ".trc"
      NextTracer.Start
      CurrentTracer.Stop

      ' Toggle current/next tracers
      If NextTracer Is tracer1 Then
        Set CurrentTracer = tracer1
        Set NextTracer = tracer2
      Else
        Set CurrentTracer = tracer2
        Set NextTracer = tracer1
      End If
    End If

  Loop

End Sub



Sub btn_DriveMode_Click()
	
	'Dim ObjPanel, ObjLED, ObjCtrl, SigLED, SigCtrl
	Dim i, SigStop, msg, TxVal
	
	
	'Get the objects on the panel
	Set ObjPanel = Documents("TESTPANEL.ipf").ActiveWindow.Object		'Selects the active Panel
	Set ObjLED = ObjPanel.ActiveScene.Controls("led_DriveMode")			'Selects the LED Indicator
	Set ObjCtrl(0) = ObjPanel.ActiveScene.Controls("ctrl_BMS_State")		'Selects the control ctrl_BMS_State
	
	Set SigLED = Signals("LED_DM")												'Selects the Signal Name for the LED Indicator
	Set SigCtrl = Signals("distributed_60B1.cmd_state.cmd_00_state_trans")		'Selects the Signal in the controller
	Set SigStop = Signals("Stop")
	
	SigStop.Value = 0
	
	'Set the LED threshold, flashing value, and value to greater than 5 for blinking LED
	ObjLED.threshold = 5
	ObjLED.Flashing = True	
	
	SigLED.Value = 10
	ObjCtrl.CycleTime = 100
	
	Set TxVal = Signals("TransmitValue")  'Created a sym file with ID, Len, CycleTime=100, and Var=TransmitValue unsigned 0,8 /max:254 	
	
	'WaNT TO SEND THE MESSAGE 5 times
	i = 0
	while (i < 5) or (SigStop.Value > 0)
	
		TxVal = "BMS_STATE_STANDBY"
		i = i + 1
		wait(100)
	wend
		
End Sub
