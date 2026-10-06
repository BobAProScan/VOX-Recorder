Public Class WinAPI

	Public Declare Function SetProcessDPIAware Lib "user32.dll" () As Boolean
	Public Declare Function CreateEvent Lib "kernel32.dll" Alias "CreateEventA" (lpEventAttributes As IntPtr, bManualReset As Boolean, bInitialState As Boolean, lpName As String) As IntPtr
	Public Declare Sub ResetEvent Lib "kernel32.dll" (hWnd As IntPtr)
	Public Declare Function WaitForSingleObject Lib "kernel32.dll" (Handle As IntPtr, MilliSeconds As Integer) As Integer
	Public Declare Function waveInOpen Lib "winmm.dll" (ByRef WaveInPtr As IntPtr, uDeviceID As Integer, ByRef lpFormat As WAVEFORMATEX, dwCallback As IntPtr, dwInstance As Integer, dwFlags As Integer) As Integer
	Public Declare Function waveInAddBuffer Lib "winmm.dll" (WaveInPtr As IntPtr, ByRef pwh As WAVEHDR, cbwh As Integer) As Integer
	Public Declare Function waveInReset Lib "winmm.dll" (WaveInPtr As IntPtr) As Integer
	Public Declare Function waveInClose Lib "winmm.dll" (WaveInPtr As IntPtr) As Integer
	Public Declare Function waveInStart Lib "winmm.dll" (WaveInPtr As IntPtr) As Integer
	Public Declare Function waveInUnprepareHeader Lib "winmm.dll" (WaveInPtr As IntPtr, ByRef pwh As WAVEHDR, cbwh As Integer) As Integer
	Public Declare Function waveInPrepareHeader Lib "winmm.dll" (WaveInPtr As IntPtr, ByRef pwh As WAVEHDR, cbwh As Integer) As Integer
	Public Declare Function waveInGetNumDevs Lib "winmm.dll" () As Integer
	Public Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (uDeviceID As Integer, ByRef lpCaps As WAVEINCAPS, uSize As Integer) As Integer
	Public Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (err As Integer, lpText As String, uSize As Integer) As Integer
	Public Declare Function beInitStream Lib "lame_enc.dll" Alias "beInitStream" (ByRef pbeConfig As BE_CONFIG_FORMAT_LHV1, ByRef pdwSamples As Integer, ByRef pdwBufferSize As Integer, ByRef hbeStream As Integer) As Integer
	Public Declare Function beEncodeChunk Lib "lame_enc.dll" Alias "beEncodeChunk" (hbeStream As Integer, nSamples As Integer, pInput As IntPtr, pOutput() As Byte, ByRef pdwOutput As Integer) As Integer
	Public Declare Function beDeinitStream Lib "lame_enc.dll" Alias "beDeinitStream" (hbeStream As Integer, pOutput() As Byte, ByRef pdwOutput As Integer) As Integer
	Public Declare Function beCloseStream Lib "lame_enc.dll" Alias "beCloseStream" (ByRef hbeStream As Integer) As Integer
	Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (lpszCommand As String, lpszReturnString As String, cchReturnLength As Integer, winHandle As IntPtr) As Integer
	Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (fdwError As Integer, lpszErrorText As String, cchErrorText As Integer) As Boolean
	Public Declare Function SHGetKnownFolderPath Lib "shell32.dll" (ByRef ID As Guid, Flags As Integer, Token As IntPtr, ByRef Path As IntPtr) As Integer

	Public Const WAVE_MAPPER As Integer = -1
	Public Const WAVE_FORMAT_PCM As Integer = 1
	Public Const WHDR_DONE As Integer = &H1
	Public Const WOM_DONE As Integer = &H3BD
	Public Const CALLBACK_EVENT As Integer = &H50000

	<StructLayoutAttribute(LayoutKind.Sequential)>
	Public Structure WAVEHDR
		Dim lpData As IntPtr
		Dim dwBufferLength As Integer
		Dim dwBytesRecorded As Integer
		Dim dwUser As IntPtr
		Dim dwFlags As Integer
		Dim dwLoops As Integer
		Dim lpNext As IntPtr
		Dim reserved As Integer
	End Structure

	<StructLayoutAttribute(LayoutKind.Sequential)>
	Public Structure WAVEFORMATEX
		Dim wFormatTag As Short
		Dim nChannels As Short
		Dim nSamplesPerSec As Integer
		Dim nAvgBytesPerSec As Integer
		Dim nBlockAlign As Short
		Dim wBitsPerSample As Short
	End Structure

	<StructLayout(LayoutKind.Sequential)>
	Public Structure WAVEINCAPS
		Dim wMid As Short
		Dim wPid As Short
		Dim vDriverVersion As Integer
		<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>
		Dim szPname As String
		Dim dwFormats As Integer
		Dim wChannels As Short
		Dim wReserved As Short
	End Structure

	<StructLayoutAttribute(LayoutKind.Sequential, Size:=331)>
	Public Structure BE_CONFIG_FORMAT_LHV1
		Dim dwConfig As Integer
		Dim dwStructVersion As Integer
		Dim dwStructSize As Integer
		Dim dwSampleRate As Integer
		Dim dwReSampleRate As Integer
		Dim nMode As Integer
		Dim dwBitrate As Integer
		Dim dwMaxBitrate As Integer
		Dim nPreset As Integer
		Dim dwMpegVersion As Integer
		Dim dwPsyModel As Integer
		Dim dwEmphasis As Integer
		Dim bPrivate As Integer
		Dim bCRC As Integer
		Dim bCopyright As Integer
		Dim bOriginal As Integer
		Dim bWriteVBRHeader As Integer
		Dim bEnableVBR As Integer
		Dim nVBRQuality As Integer
		Dim dwVbrAbr_bps As Integer
		Dim nVbrMethod As Integer
		Dim bNoRes As Integer
		Dim bStrictIso As Integer
		Dim nQuality As Integer
	End Structure

	Public Enum MMRESULT
		MMSYSERR_NOERROR = 0
		MMSYSERR_ERROR = 1
		MMSYSERR_BADDEVICEID = 2
		MMSYSERR_NOTENABLED = 3
		MMSYSERR_ALLOCATED = 4
		MMSYSERR_INVALHANDLE = 5
		MMSYSERR_NODRIVER = 6
		MMSYSERR_NOMEM = 7
		MMSYSERR_NOTSUPPORTED = 8
		MMSYSERR_BADERRNUM = 9
		MMSYSERR_INVALFLAG = 10
		MMSYSERR_INVALPARAM = 11
		MMSYSERR_HANDLEBUSY = 12
		MMSYSERR_INVALIDALIAS = 13
		MMSYSERR_BADDB = 14
		MMSYSERR_KEYNOTFOUND = 15
		MMSYSERR_READERROR = 16
		MMSYSERR_WRITEERROR = 17
		MMSYSERR_DELETEERROR = 18
		MMSYSERR_VALNOTFOUND = 19
		MMSYSERR_NODRIVERCB = 20
		WAVERR_BADFORMAT = 32
		WAVERR_STILLPLAYING = 33
		WAVERR_UNPREPARED = 34
	End Enum

	Public Shared Function GetErrorDesc(ErrorCode As Integer) As String

		Dim Message As String

		Select Case ErrorCode
			Case MMRESULT.MMSYSERR_NOERROR
				Message = String.Empty
			Case MMRESULT.MMSYSERR_ERROR
				Message = "Undefined external error."
			Case MMRESULT.MMSYSERR_BADDEVICEID
				Message = "A device ID has been used that is out of range for your system."
			Case MMRESULT.MMSYSERR_NOTENABLED
				Message = "The driver was not enabled."
			Case MMRESULT.MMSYSERR_ALLOCATED
				Message = "The specified device is already in use."
			Case MMRESULT.MMSYSERR_INVALHANDLE
				Message = "The specified device handle is invalid."
			Case MMRESULT.MMSYSERR_NODRIVER
				Message = "There is no driver installed on your system."
			Case MMRESULT.MMSYSERR_NOMEM
				Message = "There is not enough memory available for this task."
			Case MMRESULT.MMSYSERR_NOTSUPPORTED
				Message = "This function is not supported. "
			Case MMRESULT.MMSYSERR_BADERRNUM
				Message = "An error number was specified that is not defined in the system."
			Case MMRESULT.MMSYSERR_INVALFLAG
				Message = "An invalid flag was passed to a system function."
			Case MMRESULT.MMSYSERR_INVALPARAM
				Message = "An invalid parameter was passed to a system function"
			Case MMRESULT.MMSYSERR_HANDLEBUSY
				Message = "Handle being used simultaneously on another thread (eg callback)."
			Case MMRESULT.MMSYSERR_INVALIDALIAS
				Message = "Specified alias not found in WIN.INI."
			Case MMRESULT.MMSYSERR_BADDB
				Message = "The registry database is corrupt."
			Case MMRESULT.MMSYSERR_KEYNOTFOUND
				Message = "The specified registry key was not found."
			Case MMRESULT.MMSYSERR_READERROR
				Message = "The registry could not be opened or could not be read."
			Case MMRESULT.MMSYSERR_WRITEERROR
				Message = "The registry could not be written to."
			Case MMRESULT.MMSYSERR_DELETEERROR
				Message = "The specified registry key could not be deleted."
			Case MMRESULT.MMSYSERR_VALNOTFOUND
				Message = "The specified registry key value could not be found."
			Case MMRESULT.MMSYSERR_NODRIVERCB
				Message = "The driver did not generate a valid OPEN callback."
			Case MMRESULT.WAVERR_BADFORMAT
				Message = "Attempted to open WaveIn with an unsupported waveform-audio format"
			Case MMRESULT.WAVERR_STILLPLAYING
				Message = "Cannot perform this operation while media data is still playing. "
			Case MMRESULT.WAVERR_UNPREPARED
				Message = "The wave header was not prepared."
			Case Else
				Message = "Undefined error code:  " & ErrorCode
		End Select

		Return Message

	End Function

End Class
