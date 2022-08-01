// PCAN-Explorer Macro File
// First edited: 2022/8/1 19:38:05
FormatVersion=6.0

Call GCANIOInitial
Call ResetAllOutput

SubLoop: 
WaitSignal Infinite bIN_HS_VK
CheckSignal bIN_HS_VK = 1

If True
    Call EnableSUMREL	// Set SUMREL ON Immediately when recieve bIN_HS_VK

EXPOLoop:
	WaitSignal 200 bIN_HS_HK
	CheckSignal bIN_HS_HK = 1
	If True

        Call EnableSUMREL_SYNC_IN_Pulse

    	WaitSignal 40 bIN_FD_EXPO_EN	//wait for FD to output a FD_EXPO_EN within 100ms
		CheckSignal bIN_FD_EXPO_EN = 1
        If False
            GOTO SubLoop
        Call EnableSUMREL_SWR_Pulse
        wait 10
        GOTO EXPOLoop

    Call EnableSUMREL   
    GOTO SubLoop

Call ResetAllOutput
GOTO SubLoop


// Sub for CAN Messages
 
GCANIOInitial: 
                Send 1 000h 2 01h 03h
                Return

HVGInitial: 
                Send 1 000h 2 01h 03h   //TBD
                Return

ResetAllOutput: 
                Send 1 203h 1 00h
                Return            

EnableSUMREL: 
                Send 1 203h 1 01h
                Return

EnableSUMREL_SYNC_IN_Pulse: 
                Send 1 203h 1 81h	
                wait 10
                Send 1 203h 1 01h	// wait for 50ms, and sent a FD_SYNC_IN pulse of 10ms width
                Return

EnableSUMREL_SWR_Pulse: 
                Send 1 203h 1 03h				//
                wait 10
                Send 1 203h 1 01h	// IF reciever FD_EXPO_EN,  immediately sent a SWR pulse of 10ms width
                Return
                
PDOSetHVGkVmA: Send 1 000h 2 01h 03h   //TBD
                Return