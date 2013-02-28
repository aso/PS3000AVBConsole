'/**************************************************************************
'*
'* Filename:    PS3000ACSConsole.cs
'*
'* Copyright:   Pico Technology Limited 2012
'*
'* Author:      MJL
'*
'* Description:
'*   This is a console-mode program that demonstrates how to use the
'*   PS3000A driver using .NET
'*
'* Examples:
'*    Collect a block of Analogue samples immediately
'*    Collect a block of Analogue samples when a trigger event occurs
'*    Collect a stream of Analogue data 
'*    Collect a stream of Analogue data and show when a trigger event occurs
'*   
'*    Collect a block of Digital Samples immediately
'*    Collect a block of Digital Samples when a Digital trigger event occurs
'*    Collect a block of Digital and Analogue Samples when an Analogue AND a Digital trigger event occurs
'*    Collect a block of Digital and Analogue Samples when an Analogue OR a Digital trigger event occurs
'*    Collect a stream of Digital Samples 
'*    Collect a stream of Digital Samples and show Aggregated results
'*
'*
'***************************************************************************/

Imports System
Imports System.IO
Imports System.Threading
Imports PS3000AVBConsole.picotech
Imports PS3000AVBConsole.picotech.PS3000AAPI
Imports Microsoft.VisualBasic

Namespace picotech.PS3000A
    Structure ChannelSettings
        Public DCcoupled As Boolean
        Public range As PS3000AAPI.Range
        Public enabled As Boolean
    End Structure

    Structure Pwq
        Public conditions() As PS3000AAPI.PwqConditions
        Public nConditions As Short
        Public direction As PS3000AAPI.ThresholdDirection
        Public lower As UInt32
        Public upper As UInt32
        Public type As PS3000AAPI.PulseWidthType

        Public Sub New(conditions() As PS3000AAPI.PwqConditions, _
                nConditions As Short, _
                direction As PS3000AAPI.ThresholdDirection, _
                lower As UInt32, upper As UInt32, _
                type As PS3000AAPI.PulseWidthType)
            Me.conditions = conditions
            Me.nConditions = nConditions
            Me.direction = direction
            Me.lower = lower
            Me.upper = upper
            Me.type = type
        End Sub
    End Structure

    Class PS3000AConsole
        Private ReadOnly _handle As Short
        Public Const BUFFER_SIZE As Int32 = 4096    '1024
        Public Const MAX_CHANNELS As Int32 = 4
        Public Const QUAD_SCOPE As Int32 = 4
        Public Const DUAL_SCOPE As Int32 = 2

        Private _timebase As UInt32 = 8
        Private _oversample As Short = 1
        Private _scaleVoltages As Boolean = True

        Private inputRanges() As UInt32 = {10, 20, 50, 100, 200, 500, 1000, 2000, 5000, 10000, 20000, 50000}
        Private Shared _ready As Boolean = False
        Private Shared _trig As Short = 0
        Private Shared _trigAt As UInt32 = 0
        Private Shared _sampleCount As Int32 = 0
        Private Shared _startIndex As UInt32 = 0
        Private Shared _autoStop As Boolean

        Private _channelSettings() As ChannelSettings
        Private _channelCount As Int32
        Private _firstRange As PS3000AAPI.Range
        Private _lastRange As PS3000AAPI.Range
        Private _digitalPorts As Int32
        Private _callbackDelegate As PS3000AAPI.ps3000aBlockReady
        Private StreamFile As String = "stream.txt"
        Private BlockFile As String = "block.txt"
        Private _minPinned As PinnedArray(Of Short)
        Private _digiPinned As PinnedArray(Of Short)

        Private Property minPinned(i As Integer) As PinnedArray(Of Short)
            Get
                Return _minPinned
            End Get
            Set(value As PinnedArray(Of Short))
                _minPinned = value
            End Set
        End Property

        Private Property digiPinned(i As Integer) As PinnedArray(Of Short)
            Get
                Return _digiPinned
            End Get
            Set(value As PinnedArray(Of Short))
                _digiPinned = value
            End Set
        End Property

        '/****************************************************************************
        ' * Callback
        ' * used by PS4000 data streaimng collection calls, on receipt of data.
        ' * used to set global flags etc checked by user routines
        ' ****************************************************************************/
        Private Shared Sub StreamingCallback(handle As Short, _
                                 noOfSamples As Int32, _
                                 startIndex As UInt32, _
                                 ov As Short, _
                                 triggerAt As UInt32, _
                                 triggered As Short, _
                                 autoStop As Short, _
                                 pVoid As IntPtr)

            '// used for streaming
            _sampleCount = noOfSamples
            _startIndex = startIndex
            _autoStop = autoStop <> 0

            '// flag to say done reading data
            _ready = True

            '// flags to show if & where a trigger has occurred
            _trig = triggered
            _trigAt = triggerAt
        End Sub

        '/****************************************************************************
        ' * Callback
        ' * used by PS4000 data block collection calls, on receipt of data.
        ' * used to set global flags etc checked by user routines
        ' ****************************************************************************/
        Private Shared Sub BlockCallback(handle As Short, status As Short, pVoid As IntPtr)
            '// flag to say done reading data
            If status <> CShort(PS3000AAPI.PICO_CANCELLED) Then
                _ready = True
            End If
        End Sub

        '/****************************************************************************
        '* SetDefaults - restore default settings
        '****************************************************************************/
        Private Sub SetDefaults()
            Dim i As Integer

            For i = 0 To _channelCount - 1 '// reset channels to most recent settings
                PS3000AAPI.SetChannel(_handle, _
                                   PS3000AAPI.Channel.ChannelA + i, _
                                   IIf(_channelSettings(PS3000AAPI.Channel.ChannelA + i).enabled, 1, 0), _
                                   IIf(_channelSettings(PS3000AAPI.Channel.ChannelA + i).DCcoupled, 1, 0), _
                                   _channelSettings(PS3000AAPI.Channel.ChannelA + i).range, _
                                   0)
            Next
        End Sub

        '/****************************************************************************
        '* SetDigitals - enable Digital Channels
        '****************************************************************************/
        Private Sub SetDigitals()
            Dim port As PS3000AAPI.Channel
            Dim status As Short
            Dim logicLevel As Short
            Dim logicVoltage As Single = 1.5
            Dim maxLogicVoltage As Short = 5
            Dim enabled As Short = 1

            status = CShort(PS3000AAPI.PICO_OK)

            '// Set logic threshold
            logicLevel = CShort(((logicVoltage / maxLogicVoltage) * PS3000AAPI.MaxLogicLevel))

            '// Enable Digital ports
            For port = PS3000AAPI.Channel.PS3000A_DIGITAL_PORT0 To PS3000AAPI.Channel.PS3000A_DIGITAL_PORT2 - 1
                status = PS3000AAPI.SetDigitalPort(_handle, port, enabled, logicLevel)
            Next
            Console.WriteLine(IIf(status <> CShort(PS3000AAPI.PICO_OK), "SetDigitals:PS3000AAPI.SetDigitalPort Status = 0x{0:X6}", ""), status)
        End Sub


        '/****************************************************************************
        '* DisableDigital - disable Digital Channels
        '****************************************************************************/
        Private Sub DisableDigital()
            Dim port As PS3000AAPI.Channel
            Dim status As Short

            status = CShort(PS3000AAPI.PICO_OK)

            '// Disable Digital ports 
            For port = PS3000AAPI.Channel.PS3000A_DIGITAL_PORT0 To PS3000AAPI.Channel.PS3000A_DIGITAL_PORT1
                status = PS3000AAPI.SetDigitalPort(_handle, port, 0, 0)
            Next
            Console.WriteLine(IIf(status <> CShort(PS3000AAPI.PICO_OK), "DisableDigital:PS3000AAPI.SetDigitalPort Status = 0x{0:X6}", ""), status)
        End Sub

        '/****************************************************************************
        '* DisableAnalogue - disable analogue Channels
        '****************************************************************************/
        Private Sub DisableAnalogue()
            Dim status As Short
            Dim i As Integer

            status = CShort(PS3000AAPI.PICO_OK)

            '// Disable analogue ports
            For i = 0 To _channelCount - 1
                status = PS3000AAPI.SetChannel(_handle, PS3000AAPI.Channel.ChannelA + i, 0, 0, 0, 0)
            Next
            Console.WriteLine(IIf(status <> CShort(PS3000AAPI.PICO_OK), "DisableAnalogue:PS3000AAPI.SetChannel Status = 0x{0:X6}", ""), status)
        End Sub

        '/****************************************************************************
        '* adc_to_mv
        '*
        '* Convert an 16-bit ADC count into millivolts
        '****************************************************************************/
        Private Function adc_to_mv(raw As Integer, range As PS3000AAPI.Range) As Integer
            Return CInt((CInt(raw) * inputRanges(range)) / PS3000AAPI.MaxValue)
        End Function

        '/****************************************************************************
        '* mv_to_adc
        '*
        '* Convert a millivolt value into a 16-bit ADC count
        '*
        '*  (useful for setting trigger thresholds)
        '****************************************************************************/
        Private Function mv_to_adc(mv As Short, range As PS3000AAPI.Range) As Int16
            Return CType((CInt(mv) * PS3000AAPI.MaxValue) / inputRanges(range), Int16)
        End Function

        '/****************************************************************************
        '* BlockDataHandler
        '* - Used by all block data routines
        '* - acquires data (user sets trigger mode before calling), displays 10 items
        '*   and saves all to block.txt
        '* Input :
        '* - text : the text to display before the display of data slice
        '* - offset : the offset into the data buffer to start the display's slice.
        '****************************************************************************/
        Private Sub BlockDataHandler(text As String, offset As Int32, mode As PS3000AAPI.Mode)
            Dim status As Short
            Dim retry As Boolean
            Dim sampleCount As UInt32 = BUFFER_SIZE
            Dim minPinned() As PinnedArray(Of Short) = New PinnedArray(Of Short)(_channelCount) {}
            Dim maxPinned() As PinnedArray(Of Short) = New PinnedArray(Of Short)(_channelCount) {}
            Dim digiPinned() As PinnedArray(Of Short) = New PinnedArray(Of Short)(_digitalPorts) {}

            Dim timeIndisposed As Integer

            If mode = PS3000AAPI.Mode.ANALOGUE Or mode = PS3000AAPI.Mode.MIXED Then
                Dim i As Integer

                For i = 0 To _channelCount - 1
                    Dim minBuffers(sampleCount) As Short
                    Dim maxBuffers(sampleCount) As Short

                    minPinned(i) = New PinnedArray(Of Short)(minBuffers)
                    maxPinned(i) = New PinnedArray(Of Short)(maxBuffers)
                    status = PS3000AAPI.SetDataBuffers(_handle, CType(i, PS3000AAPI.Channel), maxBuffers, minBuffers, CInt(sampleCount), 0, PS3000AAPI.RatioMode.None)
                    Console.WriteLine(IIf(status <> CShort(PS3000AAPI.PICO_OK), "BlockDataHandler:PS3000AAPI.SetDataBuffers Channel {0} Status = 0x{1:X6}", ""), Chr(Asc("A") + i), status)
                Next
            End If

            If mode = PS3000AAPI.Mode.DIGITAL Or mode = PS3000AAPI.Mode.MIXED Then
                Dim i As Integer
                For i = 0 To _digitalPorts - 1
                    Dim digiBuffer(sampleCount) As Short

                    digiPinned(i) = New PinnedArray(Of Short)(digiBuffer)
                    status = PS3000AAPI.SetDataBuffer(_handle, i + PS3000AAPI.Channel.PS3000A_DIGITAL_PORT0, digiBuffer, CInt(sampleCount), 0, PS3000AAPI.RatioMode.None)
                    Console.WriteLine(IIf(status <> CShort(PS3000AAPI.PICO_OK), "BlockDataHandler:PS3000AAPI.SetDataBuffer {0} Status = 0x{1,0:X6}", ""), i + PS3000AAPI.Channel.PS3000A_DIGITAL_PORT0, status)
                Next
            End If

            '/*  find the maximum number of samples, the time interval (in timeUnits),
            '*		 the most suitable time units, and the maximum _oversample at the current _timebase*/
            Dim timeInterval As Integer
            Dim maxSamples As Integer

            While (PS3000AAPI.GetTimebase(_handle, _timebase, CInt(sampleCount), timeInterval, _oversample, maxSamples, 0) <> 0)
                Console.WriteLine("Selected timebase {0} could not be used", _timebase)
                _timebase += 1
            End While
            Console.Write("Timebase: {0}", _timebase)
            Console.WriteLine(vbTab & "oversample:{0}", _oversample)

            '/* Start it collecting, then wait for completion*/
            _ready = False
            _callbackDelegate = AddressOf BlockCallback

            Do
                retry = False
                status = PS3000AAPI.RunBlock(_handle, 0, CInt(sampleCount), _timebase, _oversample, timeIndisposed, 0, _callbackDelegate, IntPtr.Zero)
                If status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_CONNECTED) Or status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_NOT_CONNECTED) Or status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_UNDERVOLTAGE) Then
                    status = PowerSourceSwitch(_handle, status)
                    retry = True
                Else
                    Console.WriteLine(IIf(status <> CShort(PS3000AAPI.PICO_OK), "BlockDataHandler:PS3000AAPI.RunBlock Status = 0x{0:X6}", ""), status)
                End If
            Loop While retry

            Console.WriteLine("Waiting for data...Press a key to abort")

            While Not _ready And Not Console.KeyAvailable
                Thread.Sleep(100)
            End While

            If Console.KeyAvailable Then
                Console.ReadKey(True) ' // clear the key
            End If

            PS3000AAPI.StopUnit(_handle)

            If _ready Then
                Dim overflow As Short

                status = PS3000AAPI.GetValues(_handle, 0, sampleCount, 1, PS3000AAPI.RatioMode.None, 0, overflow)
                If status = CShort(PS3000AAPI.PICO_OK) Then
                    '/* Print out the first 10 readings, converting the readings to mV if required */
                    Console.WriteLine(text)

                    Dim ch As Integer

                    If mode = PS3000AAPI.Mode.ANALOGUE Or mode = PS3000AAPI.Mode.MIXED Then
                        Console.WriteLine("Value {0}", IIf(_scaleVoltages, "mV", "ADC Counts"))
                        For ch = 0 To _channelCount - 1
                            If _channelSettings(ch).enabled Then
                                Console.Write("Channel{0}          ", Chr(Asc("A") + ch))
                            End If
                        Next
                    End If

                    If mode = PS3000AAPI.Mode.DIGITAL Or mode = PS3000AAPI.Mode.MIXED Then
                        Console.Write("DIGITAL VALUE")
                    End If

                    Console.WriteLine()

                    For i = offset To offset + 10 - 1
                        If mode = PS3000AAPI.Mode.ANALOGUE Or mode = PS3000AAPI.Mode.MIXED Then
                            For ch = 0 To _channelCount - 1
                                If _channelSettings(ch).enabled Then
                                    '// If _scaleVoltages, show mV values
                                    '// else show ADC counts
                                    Debug.Print(maxPinned(ch).Target(i))
                                    Debug.Print(CInt(_channelSettings(CInt(PS3000AAPI.Channel.ChannelA + ch)).range))
                                    Debug.Print(adc_to_mv(maxPinned(ch).Target(i), _channelSettings(PS3000AAPI.Channel.ChannelA + ch).range))

                                    Console.Write("{0,8}          ", _
                                        IIf(_scaleVoltages, adc_to_mv(maxPinned(ch).Target(i), _
                                             _channelSettings(PS3000AAPI.Channel.ChannelA + ch).range), _
                                            maxPinned(ch).Target(i) _
                                        ) _
                                    )
                                End If
                            Next
                        End If

                        If mode = PS3000AAPI.Mode.DIGITAL Or mode = PS3000AAPI.Mode.MIXED Then
                            Dim digiValue As Short = digiPinned(1).Target(i)
                            digiValue <<= 8
                            digiValue = digiValue Or digiPinned(0).Target(i)
                            Console.Write("0x{0,4:X}", digiValue.ToString("X4"))
                        End If

                        Console.WriteLine()
                    Next

                    If mode = PS3000AAPI.Mode.ANALOGUE Or mode = PS3000AAPI.Mode.MIXED Then
                        sampleCount = Math.Min(sampleCount, BUFFER_SIZE)
                        Dim writer As TextWriter = New StreamWriter(BlockFile, False)
                        writer.Write("For each of the {0} Channels, results shown are....", _channelCount)
                        writer.WriteLine()
                        writer.WriteLine("Time interval Maximum Aggregated value ADC Count & mV, Minimum Aggregated value ADC Count & mV")
                        writer.WriteLine()

                        For i = 0 To _channelCount - 1
                            If _channelSettings(i).enabled Then
                                writer.Write("Time   Ch  Max ADC    Max mV   Min ADC    Min mV   ")
                            End If
                        Next

                        writer.WriteLine()

                        For i = 0 To sampleCount - 1
                            For ch = 0 To _channelCount - 1
                                If _channelSettings(ch).enabled Then
                                    writer.Write("{0,5}  ", (i * timeInterval))
                                    writer.Write( _
                                        "Ch{0} {1,7}   {2,7}   {3,7}   {4,7}   ", _
                                        Chr(Asc("A") + ch), _
                                        maxPinned(ch).Target(i), _
                                        adc_to_mv(maxPinned(ch).Target(i), _channelSettings(PS3000AAPI.Channel.ChannelA + ch).range), _
                                        minPinned(ch).Target(i), _
                                        adc_to_mv(minPinned(ch).Target(i), _channelSettings(PS3000AAPI.Channel.ChannelA + ch).range) _
                                        )
                                End If
                            Next
                            writer.WriteLine()
                        Next
                        writer.Close()
                    End If
                Else
                    If status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_CONNECTED) Or status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_NOT_CONNECTED) Or status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_UNDERVOLTAGE) Then
                        If status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_UNDERVOLTAGE) Then
                            status = PowerSourceSwitch(_handle, status)
                        Else
                            Console.WriteLine("Power source changed. Data collection aborted")
                        End If
                    Else
                        Console.WriteLine("BlockDataHandler:PS3000AAPI.GetValues Status = 0x{0:X6}", status)
                    End If
                End If
            Else
                Console.WriteLine("data collection aborted")
            End If

            PS3000AAPI.StopUnit(_handle)

            If mode = PS3000AAPI.Mode.ANALOGUE Or mode = PS3000AAPI.Mode.MIXED Then
                For Each p As PinnedArray(Of Short) In minPinned
                    If Not IsNothing(p) Then
                        p.Dispose()
                    End If
                Next

                For Each p As PinnedArray(Of Short) In maxPinned
                    If Not IsNothing(p) Then
                        p.Dispose()
                    End If
                Next
            End If

            If mode = PS3000AAPI.Mode.DIGITAL Or mode = PS3000AAPI.Mode.MIXED Then
                For Each p As PinnedArray(Of Short) In digiPinned
                    If Not IsNothing(p) Then
                        p.Dispose()
                    End If
                Next
            End If

        End Sub

        '/****************************************************************************
        ' * RapidBlockDataHandler
        ' * - Used by all the CollectBlockRapid routine
        ' * - acquires data (user sets trigger mode before calling), displays 10 items
        ' *   and saves all to data.txt
        ' * Input :
        ' * - nRapidCaptures : the user specified number of blocks to capture
        ' ****************************************************************************/
        Private Sub RapidBlockDataHandler(nRapidCaptures As UShort)
            Dim status As Short
            Dim numChannels As Integer = _channelCount
            Dim numSamples As UInt32 = BUFFER_SIZE

            '// Run the rapid block capture
            Dim timeIndisposed As Int32
            _ready = False

            _callbackDelegate = AddressOf BlockCallback
            PS3000AAPI.RunBlock(_handle,
                        0,
                        CInt(numSamples),
                        _timebase,
                        _oversample,
                        timeIndisposed,
                        0,
                        _callbackDelegate,
                        IntPtr.Zero)

            Console.WriteLine("Waiting for data...Press a key to abort")

            While Not _ready And Not Console.KeyAvailable
                Thread.Sleep(100)
            End While
            If Console.KeyAvailable Then
                Console.ReadKey(True) '// clear the key
            End If

            PS3000AAPI.StopUnit(_handle)

            '// Set up the data arrays and pin them
            Dim values As Short()()() = New Short(nRapidCaptures)()() {}
            Dim pinned(nRapidCaptures, numChannels) As PinnedArray(Of Short)

            For segment As UShort = 0 To nRapidCaptures - 1
                values(segment) = New Short(numChannels)() {}
                For channel As Short = 0 To numChannels - 1
                    If _channelSettings(channel).enabled Then
                        ReDim values(segment)(channel)(numSamples)
                        pinned(segment, channel) = New PinnedArray(Of Short)(values(segment)(channel))

                        status = PS3000AAPI.SetDataBuffer(_handle,
                                               CType(channel, PS3000AAPI.Channel),
                                               values(segment)(channel),
                                               CInt(numSamples),
                                               segment,
                                               0)
                        Console.WriteLine(IIf(status <> CShort(PS3000AAPI.PICO_OK), "RapidBlockDataHandler:PS3000AAPI.SetDataBuffer Channel {0} Status = 0x{1:X6}", ""), Chr(Asc("A") + channel), status)

                    Else
                        status = PS3000AAPI.SetDataBuffer(_handle, _
                                   CType(channel, PS3000AAPI.Channel), _
                                    Nothing, _
                                    0, _
                                    segment, _
                                    0)
                        Console.WriteLine(IIf(status <> CShort(PS3000AAPI.PICO_OK), "RapidBlockDataHandler:PS3000AAPI.SetDataBuffer Channel {0} Status = 0x{1:X6}", ""), Chr(Asc("A") + channel), status)
                    End If
                Next
            Next

            '// Read the data
            Dim overflows As Short() = New Short(nRapidCaptures) {}

            status = PS3000AAPI.GetValuesRapid(_handle, numSamples, 0, CUShort(nRapidCaptures - 1), 1, PS3000AAPI.RatioMode.None, overflows)

            '/* Print out the first 10 readings, converting the readings to mV if required */
            Console.WriteLine("Values in {0}", IIf(_scaleVoltages, "mV", "ADC Counts"))

            For seg As Integer = 0 To nRapidCaptures - 1
                Console.WriteLine("Capture {0}", seg)
                For i As Integer = 0 To 10 - 1
                    For chan As Integer = 0 To _channelCount - 1
                        If (_channelSettings(chan).enabled) Then
                            Console.Write( _
                                "{0}" & vbTab, _
                                IIf( _
                                    _scaleVoltages, _
                                    adc_to_mv( _
                                        pinned(seg, chan).Target(i), _
                                        _channelSettings(PS3000AAPI.Channel.ChannelA + chan).range _
                                    ), _
                                    pinned(seg, chan).Target(i) _
                                ) _
                            )

                            '// If _scaleVoltages, show mV values
                            '// else show ADC counts
                        End If
                    Next
                    Console.WriteLine()
                Next
                Console.WriteLine()
            Next

            ' // Un-pin the arrays
            For Each p As PinnedArray(Of Short) In pinned
                If Not IsNothing(p) Then
                    p.Dispose()
                End If
            Next

            ' //TODO: Do what ever is required with the data here.
        End Sub

        '/****************************************************************************
        '*  SetTrigger  (Non-Digital Version)
        '*  this function sets all the required trigger parameters, and calls the 
        '*  triggering functions
        '****************************************************************************/
        Private Function SetTrigger(channelProperties() As PS3000AAPI.TriggerChannelProperties, _
                nChannelProperties As Short, _
                triggerConditions() As PS3000AAPI.TriggerConditions, _
                nTriggerConditions As Short, _
                directions() As PS3000AAPI.ThresholdDirection, _
                pwq As Pwq, _
                delay As UInt32, _
                auxOutputEnabled As Short, _
                autoTriggerMs As Int32) As Short

            Dim status As Short

            If (status = PS3000AAPI.SetTriggerChannelProperties(_handle, channelProperties, nChannelProperties, auxOutputEnabled, autoTriggerMs)) <> 0 Then
                Return status
            End If

            If (status = PS3000AAPI.SetTriggerChannelConditions(_handle, triggerConditions, nTriggerConditions)) <> 0 Then
                Return status
            End If

            If IsNothing(directions) Then
                directions = New PS3000AAPI.ThresholdDirection() {
                    PS3000AAPI.ThresholdDirection.None,
                    PS3000AAPI.ThresholdDirection.None,
                    PS3000AAPI.ThresholdDirection.None,
                    PS3000AAPI.ThresholdDirection.None,
                    PS3000AAPI.ThresholdDirection.None,
                    PS3000AAPI.ThresholdDirection.None
                }
            End If

            status = PS3000AAPI.SetTriggerChannelDirections(_handle, _
                directions(CInt(PS3000AAPI.Channel.ChannelA)), _
                directions(CInt(PS3000AAPI.Channel.ChannelB)), _
                directions(CInt(PS3000AAPI.Channel.ChannelC)), _
                directions(CInt(PS3000AAPI.Channel.ChannelD)), _
                directions(CInt(PS3000AAPI.Channel.External)), _
                directions(CInt(PS3000AAPI.Channel.Aux)))

            If status <> 0 Then
                Return status
            End If

            status = PS3000AAPI.SetTriggerDelay(_handle, delay)
            If status <> 0 Then
                Return status
            End If

            If IsNothing(pwq) Then
                pwq = New Pwq(Nothing, 0, PS3000AAPI.ThresholdDirection.None, 0, 0, PS3000AAPI.PulseWidthType.None)
            End If

            status = PS3000AAPI.SetPulseWidthQualifier(_handle, pwq.conditions,
                                                    pwq.nConditions, pwq.direction,
                                                    pwq.lower, pwq.upper, pwq.type)

            Return status
        End Function

        '/****************************************************************************
        '*  SetTrigger
        '*  this overloaded version of SetTrigger includes digital parameters
        '****************************************************************************/
        Private Function SetTrigger(channelProperties() As PS3000AAPI.TriggerChannelProperties, _
                nChannelProperties As Short, _
                triggerConditions() As PS3000AAPI.TriggerConditionsV2, _
                nTriggerConditions As Short, _
                directions() As PS3000AAPI.ThresholdDirection, _
                pwq As Pwq, _
                delay As UInt32, _
                auxOutputEnabled As Short, _
                autoTriggerMs As Int32, _
                digitalDirections() As PS3000AAPI.DigitalChannelDirections, _
                nDigitalDirections As Short) As Short
            Dim status As Short

            If (status = PS3000AAPI.SetTriggerChannelProperties(_handle, channelProperties, nChannelProperties, auxOutputEnabled, autoTriggerMs)) <> 0 Then
                Return status
            End If

            If (status = PS3000AAPI.SetTriggerChannelConditionsV2(_handle, triggerConditions, nTriggerConditions)) <> 0 Then
                Return status
            End If

            If IsNothing(directions) Then
                directions = New PS3000AAPI.ThresholdDirection() {
                    PS3000AAPI.ThresholdDirection.None,
                    PS3000AAPI.ThresholdDirection.None,
                    PS3000AAPI.ThresholdDirection.None,
                    PS3000AAPI.ThresholdDirection.None,
                    PS3000AAPI.ThresholdDirection.None,
                    PS3000AAPI.ThresholdDirection.None
                }
            End If

            If ((status = PS3000AAPI.SetTriggerChannelDirections(_handle,
                    directions(CInt(PS3000AAPI.Channel.ChannelA)),
                    directions(CInt(PS3000AAPI.Channel.ChannelB)),
                    directions(CInt(PS3000AAPI.Channel.ChannelC)),
                    directions(CInt(PS3000AAPI.Channel.ChannelD)),
                    directions(CInt(PS3000AAPI.Channel.External)),
                    directions(CInt(PS3000AAPI.Channel.Aux)))) <> 0) Then

                Return status
            End If

            If (status = PS3000AAPI.SetTriggerDelay(_handle, delay)) <> 0 Then
                Return status
            End If

            If IsNothing(pwq) Then
                pwq = New Pwq(Nothing, 0, PS3000AAPI.ThresholdDirection.None, 0, 0, PS3000AAPI.PulseWidthType.None)
            End If

            status = PS3000AAPI.SetPulseWidthQualifier(_handle, _
                pwq.conditions, _
                pwq.nConditions, _
                pwq.direction, _
                pwq.lower, _
                pwq.upper, _
                pwq.type)

            If (_digitalPorts > 0) Then
                If ((status = PS3000AAPI.SetTriggerDigitalPort(_handle, digitalDirections, nDigitalDirections)) <> 0) Then
                    Return status
                End If
            End If

            Return status
        End Function

        '/****************************************************************************
        '* CollectBlockImmediate
        '*  this function demonstrates how to collect a single block of data
        '*  from the unit (start collecting immediately)
        '****************************************************************************/
        Private Sub CollectBlockImmediate()
            Console.WriteLine("Collect block immediate...")
            Console.WriteLine("Press a key to start")
            WaitForKey()
            SetDefaults()

            '/* Trigger disabled	*/
            SetTrigger({}, 0, {}, 0, {}, Nothing, 0, 0, 0)
            BlockDataHandler("First 10 readings", 0, PS3000AAPI.Mode.ANALOGUE)
        End Sub

        '/****************************************************************************
        '*  CollectBlockRapid
        '*  this function demonstrates how to collect blocks of data
        '* using the RapidCapture function
        '****************************************************************************/
        Private Sub CollectBlockRapid()
            Dim numRapidCaptures As UShort = 1
            Dim valid As Boolean = False

            Console.WriteLine("Collect rapid block...")
            Console.WriteLine("Specify number of captures:")

            Do
                Try
                    numRapidCaptures = UShort.Parse(Console.ReadLine())
                    valid = True
                Catch ex As Exception
                    valid = False
                    Console.WriteLine("Enter numeric values only")
                End Try
            Loop While (PS3000AAPI.SetNoOfRapidCaptures(_handle, numRapidCaptures) > 0 Or Not valid)

            Dim maxSamples As Integer
            PS3000AAPI.MemorySegments(_handle, numRapidCaptures, maxSamples)

            Console.WriteLine("Collecting {0} rapid blocks. Press a key to start", numRapidCaptures)

            WaitForKey()
            SetDefaults()

            ' /* Trigger is optional, disable it for now	*/
            SetTrigger({}, 0, {}, 0, {}, Nothing, 0, 0, 0)

            RapidBlockDataHandler(numRapidCaptures)
        End Sub

        ' /****************************************************************************
        '* CollectBlockTriggered
        '*  this function demonstrates how to collect a single block of data from the
        '*  unit, when a trigger event occurs.
        '****************************************************************************/
        Private Sub CollectBlockTriggered()
            Dim triggerVoltage As Short = mv_to_adc(inputRanges(_channelSettings(PS3000AAPI.Channel.ChannelA).range) / 5, _channelSettings(PS3000AAPI.Channel.ChannelA).range) '// ChannelInfo stores ADC counts
            Dim sourceDetails() As PS3000AAPI.TriggerChannelProperties = New PS3000AAPI.TriggerChannelProperties() {
                New PS3000AAPI.TriggerChannelProperties(triggerVoltage, _
                256 * 10, _
                triggerVoltage, _
                256 * 10, _
                PS3000AAPI.Channel.ChannelA, _
                PS3000AAPI.ThresholdMode.Level)
            }

            Dim conditions() As PS3000AAPI.TriggerConditions = New PS3000AAPI.TriggerConditions() {
                New PS3000AAPI.TriggerConditions( _
                    PS3000AAPI.TriggerState.TrigTrue, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare)
            }

            Dim directions() As PS3000AAPI.ThresholdDirection = New PS3000AAPI.ThresholdDirection() {
                PS3000AAPI.ThresholdDirection.Rising,
                PS3000AAPI.ThresholdDirection.None,
                PS3000AAPI.ThresholdDirection.None,
                PS3000AAPI.ThresholdDirection.None,
                PS3000AAPI.ThresholdDirection.None,
                PS3000AAPI.ThresholdDirection.None
            }

            Console.WriteLine("Collect block triggered...")


            Console.Write("Collects when value rises past {0}", IIf(_scaleVoltages, _
                adc_to_mv(sourceDetails(0).ThresholdMajor, _
                CInt(_channelSettings(CInt(PS3000AAPI.Channel.ChannelA)).range)) _
                , sourceDetails(0).ThresholdMajor) _
            )
            Console.WriteLine("{0}", IIf(_scaleVoltages, "mV", " ADC Counts"))

            Console.WriteLine("Press a key to start...")
            WaitForKey()

            SetDefaults()

            '/* Trigger enabled
            ' * Rising edge
            ' * Threshold = 1000mV */
            SetTrigger(sourceDetails, 1, conditions, 1, directions, Nothing, 0, 0, 0)

            BlockDataHandler("Ten readings after trigger", 0, PS3000AAPI.Mode.ANALOGUE)
        End Sub

        '/****************************************************************************
        '* Initialise unit' structure with Variant specific defaults
        '****************************************************************************/
        Private Sub GetDeviceInfo()
            Dim description() As String = {
                "Driver Version    ",
                "USB Version       ",
                "Hardware Version  ",
                "Variant Info      ",
                "Serial            ",
                "Cal Date          ",
                "Kernel Ver        ",
                "Digital Hardware  ",
                "Analogue Hardware "
            }
            Dim line As System.Text.StringBuilder = New System.Text.StringBuilder(80)

            If (_handle >= 0) Then
                For i As Integer = 0 To 9 - 1
                    Dim requiredSize As Short
                    PS3000AAPI.GetUnitInfo(_handle, line, 80, requiredSize, i)
                    If i = 3 Then
                        If line.ToString().Contains("340") Then    '// PS340XA/B device
                            _channelCount = QUAD_SCOPE
                        Else
                            _channelCount = DUAL_SCOPE
                        End If
                    End If

                    If i = 3 Then
                        If line.ToString().EndsWith("MSO") Then
                            _digitalPorts = 2
                        Else
                            _digitalPorts = 0
                        End If
                    End If
                    Console.WriteLine("{0}: {1}", description(i), line)
                Next
            End If
        End Sub

        '/****************************************************************************
        ' * Select input voltage ranges for channels A and B
        ' ****************************************************************************/
        Private Sub SetVoltages()
            Dim valid As Boolean = False
            Dim count As Short = 0
            Dim i As Integer

            ' /* See what ranges are available... */
            For i = CInt(_firstRange) To CInt(_lastRange)
                Console.WriteLine("{0} . {1} mV", i, inputRanges(i))
            Next

            Do
                '/* Ask the user to select a range */
                Console.WriteLine("Specify voltage range ({0}..{1})", _firstRange, _lastRange)
                Console.WriteLine("99 - switches channel off")

                Dim ch As Integer

                For ch = 0 To _channelCount - 1
                    Console.WriteLine("")
                    Dim range As UInt32 = 8

                    Do
                        Try
                            Console.WriteLine("Channel: {0}", Chr(Asc("A") + ch))
                            range = UInteger.Parse(Console.ReadLine())
                            valid = True
                        Catch ex As Exception
                            valid = False
                            Console.WriteLine("Enter numeric values only")
                        End Try
                    Loop While ( _
                        (range <> 99 _
                         And (range < CType(_firstRange, UInt32) Or range > CType(_lastRange, UInt32)) _
                         Or Not valid) _
                     )

                    If range <> 99 Then
                        _channelSettings(ch).range = CType(range, PS3000AAPI.Range)
                        Console.WriteLine(" = {0} mV", inputRanges(range))
                        _channelSettings(ch).enabled = True
                        count += 1
                    Else
                        Console.WriteLine("Channel Switched off")
                        _channelSettings(ch).enabled = False
                        _channelSettings(ch).range = PS3000AAPI.Range.Range_MAX_RANGE - 1
                    End If
                Next
                Console.Write(IIf(count = 0, "*** At least 1 channel must be enabled *** ", ""))
            Loop While (count = 0) '// must have at least one channel enabled

            SetDefaults()  '// Set defaults now, so that if all but 1 channels get switched off, timebase updates to timebase 0 will work
        End Sub

        '/****************************************************************************
        ' *
        ' * Select _timebase, set _oversample to on and time units as nano seconds
        ' *
        ' ****************************************************************************/
        Private Sub SetTimebase()
            Dim timeInterval As Integer
            Dim maxSamples As Integer
            Dim valid As Boolean = False

            Console.WriteLine("Specify timebase")

            Do
                Try
                    _timebase = UInteger.Parse(Console.ReadLine())
                    valid = True
                Catch ex As Exception
                    valid = False
                    Console.WriteLine("Enter numeric values only")
                End Try
            Loop While (Not valid)

            While (PS3000AAPI.GetTimebase(_handle, _timebase, BUFFER_SIZE, timeInterval, 1, maxSamples, 0) <> 0)
                Console.WriteLine("Selected timebase {0} could not be used", _timebase)
                _timebase += 1
            End While

            Console.WriteLine("Timebase {0} - {1} ns", _timebase, timeInterval)
            _oversample = 1
        End Sub

        '/****************************************************************************
        '* Stream Data Handler
        '* - Used by the two stream data examples - untriggered and triggered
        '* Inputs:
        '* - preTrigger - the number of samples in the pre-trigger phase 
        '*					(0 if no trigger has been set)
        '*	mode  - ANALOGUE, 
        '***************************************************************************/
        Private Sub StreamDataHandler(preTrigger As UInt32, mode As PS3000AAPI.Mode)
            Dim sampleCount As UInteger = BUFFER_SIZE * 10 '/*  *10 is to make sure buffer large enough */
            Dim minBuffers()() As Short = New Short(_channelCount)() {}
            Dim maxBuffers()() As Short = New Short(_channelCount)() {}
            Dim minPinned() As PinnedArray(Of Short) = New PinnedArray(Of Short)(_channelCount) {}
            Dim maxPinned() As PinnedArray(Of Short) = New PinnedArray(Of Short)(_channelCount) {}

            Dim digiBuffersA()() As Short = New Short(_digitalPorts)() {}
            Dim digiBuffersB()() As Short = New Short(_digitalPorts)() {}
            Dim digiPinnedA() As PinnedArray(Of Short) = New PinnedArray(Of Short)(_digitalPorts) {}
            Dim digiPinnedB() As PinnedArray(Of Short) = New PinnedArray(Of Short)(_digitalPorts) {}

            Dim totalsamples As UInt32 = 0
            Dim triggeredAt As UInt32 = 0
            Dim status As Short

            Dim downsampleRatio As UInt32
            Dim timeUnits As PS3000AAPI.ReportedTimeUnits
            Dim sampleInterval As UInt32
            Dim ratioMode As PS3000AAPI.RatioMode
            Dim postTrigger As UInt32
            Dim autoStop As Boolean
            Dim retry As Boolean
            Dim powerChange As Boolean = False

            sampleInterval = 1
            status = CShort(PS3000AAPI.PICO_OK)

            Select Case mode
                Case PS3000AAPI.Mode.ANALOGUE
                    For channel As Integer = 0 To _channelCount - 1 '// create data buffers
                        minBuffers(channel) = New Short(sampleCount) {}
                        maxBuffers(channel) = New Short(sampleCount) {}
                        minPinned(channel) = New PinnedArray(Of Short)(minBuffers(channel))
                        maxPinned(channel) = New PinnedArray(Of Short)(maxBuffers(channel))
                        status = PS3000AAPI.SetDataBuffers(_handle, CType(channel, PS3000AAPI.Channel), minBuffers(channel), maxBuffers(channel), CInt(sampleCount), 0, PS3000AAPI.RatioMode.Aggregate)
                        Console.Write(IIf(status <> CShort(PS3000AAPI.PICO_OK), "StreamDataHandler:PS3000AAPI.SetDataBuffers Channel {0} Status = 0x{1:X6}", ""), Chr(Asc("A") + channel), status)
                    Next

                    downsampleRatio = 1000
                    timeUnits = PS3000AAPI.ReportedTimeUnits.MicroSeconds
                    sampleInterval = 1
                    ratioMode = PS3000AAPI.RatioMode.Aggregate
                    postTrigger = 1000000
                    autoStop = True
                    sampleInterval = 1

                Case PS3000AAPI.Mode.AGGREGATED
                    For port As Integer = 0 To _digitalPorts - 1 '// create data buffers
                        digiBuffersA(port) = New Short(sampleCount) {}
                        digiBuffersB(port) = New Short(sampleCount) {}
                        digiPinnedA(port) = New PinnedArray(Of Short)(digiBuffersA(port))
                        digiPinnedB(port) = New PinnedArray(Of Short)(digiBuffersB(port))
                        status = PS3000AAPI.SetDataBuffers(_handle, CType(port, PS3000AAPI.Channel) + CShort(PS3000AAPI.Channel.PS3000A_DIGITAL_PORT0), digiBuffersA(port), digiBuffersB(port), CInt(sampleCount), 0, PS3000AAPI.RatioMode.Aggregate)
                        Console.WriteLine(IIf(status <> CShort(PS3000AAPI.PICO_OK), "StreamDataHandler:PS3000AAPI.SetDataBuffers {0} Status = 0x{1:X6}", ""), CType(port, PS3000AAPI.Channel) + CShort(PS3000AAPI.Channel.PS3000A_DIGITAL_PORT0), status)
                    Next

                    downsampleRatio = 10
                    timeUnits = PS3000AAPI.ReportedTimeUnits.MilliSeconds
                    sampleInterval = 10
                    ratioMode = PS3000AAPI.RatioMode.Aggregate
                    postTrigger = 10
                    autoStop = False

                Case PS3000AAPI.Mode.DIGITAL
                    For port As Integer = 0 To _digitalPorts - 1 '// create data buffers
                        digiBuffersA(port) = New Short(sampleCount) {}
                        digiPinnedA(port) = New PinnedArray(Of Short)(digiBuffersA(port))
                        status = PS3000AAPI.SetDataBuffer(_handle, CType(port + CShort(PS3000AAPI.Channel.PS3000A_DIGITAL_PORT0), PS3000AAPI.Channel), digiBuffersA(port), CInt(sampleCount), 0, PS3000AAPI.RatioMode.None)
                        Console.WriteLine(IIf(status <> CShort(PS3000AAPI.PICO_OK), "StreamDataHandler:PS3000AAPI.SetDataBuffer {0} Status = 0x{1:X6}", ""), CType(port + CShort(PS3000AAPI.Channel.PS3000A_DIGITAL_PORT0), PS3000AAPI.Channel), status)
                    Next
                    downsampleRatio = 1
                    timeUnits = PS3000AAPI.ReportedTimeUnits.MilliSeconds
                    sampleInterval = 10
                    ratioMode = PS3000AAPI.RatioMode.None
                    postTrigger = 10
                    autoStop = False

                Case Else
                    downsampleRatio = 1
                    timeUnits = PS3000AAPI.ReportedTimeUnits.MilliSeconds
                    sampleInterval = 10
                    ratioMode = PS3000AAPI.RatioMode.None
                    postTrigger = 10
                    autoStop = False
            End Select

            Console.WriteLine("Waiting for trigger...Press a key to abort")
            _autoStop = False

            Do
                retry = False
                status = PS3000AAPI.RunStreaming(_handle, sampleInterval, timeUnits, preTrigger, postTrigger - preTrigger, autoStop, downsampleRatio, ratioMode, sampleCount)

                If status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_CONNECTED) Or status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_NOT_CONNECTED) Or status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_UNDERVOLTAGE) Then
                    status = PowerSourceSwitch(_handle, status)
                    retry = True
                End If
            Loop While (retry)

            Console.WriteLine(IIf(status <> CShort(PS3000AAPI.PICO_OK), "StreamDataHandler:PS3000AAPI.RunStreaming Status = 0x{0,0:X6}", ""), status)
            Console.WriteLine("Streaming data...Press a key to abort")

            Dim writer As TextWriter

            If mode = PS3000AAPI.Mode.ANALOGUE Then
                writer = New StreamWriter(StreamFile, False)
            End If

            If mode = PS3000AAPI.Mode.ANALOGUE Then
                writer.WriteLine("For each of the {0} Channels, results shown are....", _channelCount)
                writer.WriteLine("Maximum Aggregated value ADC Count & mV, Minimum Aggregated value ADC Count & mV")
                writer.WriteLine()

                Dim i As Integer

                For i = 0 To _channelCount - 1
                    If _channelSettings(i).enabled Then
                        writer.Write("Ch  Max ADC    Max mV   Min ADC    Min mV   ")
                    End If
                Next
                writer.WriteLine()
            End If

            While (Not _autoStop And Not Console.KeyAvailable And Not powerChange)
                '/* Poll until data is received. Until then, GetStreamingLatestValues wont call the callback */
                Thread.Sleep(100)
                _ready = False
                status = PS3000AAPI.GetStreamingLatestValues(_handle, AddressOf StreamingCallback, IntPtr.Zero)

                If status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_CONNECTED) Or status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_NOT_CONNECTED) Or status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_UNDERVOLTAGE) Then
                    If status = CShort(PS3000AAPI.PICO_POWER_SUPPLY_UNDERVOLTAGE) Then
                        status = PowerSourceSwitch(_handle, status)
                    Else
                        Console.WriteLine("Power source changed.")
                    End If
                    powerChange = True
                End If

                If _ready And _sampleCount > 0 Then '/* can be ready and have no data, if autoStop has fired */
                    If _trig > 0 Then
                        triggeredAt = totalsamples + _trigAt
                        totalsamples += CType(_sampleCount, UInt32)

                        Console.Write("Collected {0,4} samples, index = {1,5} Total = {2,5}", _sampleCount, _startIndex, totalsamples)

                        If _trig > 0 Then
                            Console.Write(vbTab & "Trig at Index {0}", triggeredAt)

                            Console.WriteLine()
                            Dim i As UInt32

                            For i = _startIndex To (_startIndex + _sampleCount) - 1
                                If mode = PS3000AAPI.Mode.ANALOGUE Then
                                    Dim ch As Integer
                                    For ch = 0 To _channelCount - 1
                                        If _channelSettings(ch).enabled Then
                                            writer.Write("Ch{0} {1,7}   {2,7}   {3,7}   {4,7}   ", _
                                                Chr(Asc("A") + ch), _
                                                minPinned(ch).Target(i), _
                                                adc_to_mv(minPinned(ch).Target(i), _channelSettings(PS3000AAPI.Channel.ChannelA + ch).range), _
                                                maxPinned(ch).Target(i), _
                                                adc_to_mv(maxPinned(ch).Target(i), _channelSettings(PS3000AAPI.Channel.ChannelA + ch).range))
                                        End If
                                    Next
                                    writer.WriteLine()
                                End If

                                If mode = PS3000AAPI.Mode.DIGITAL Then
                                    Dim digiValue As Short = CShort(&HFF And digiPinnedA(1).Target(i))
                                    digiValue <<= 8
                                    digiValue = digiValue Or CShort(&HFF And digiPinnedA(0).Target(i))
                                    Console.Write("Index={0,4:D}: Value = 0x{1,4:X}  =  ", i, digiValue.ToString("X4"))

                                    Dim bit As Short

                                    For bit = 0 To 16 - 1
                                        Console.Write(IIf(((&H8000 >> bit) And digiValue) <> 0, "1 ", "0 "))
                                    Next
                                    Console.WriteLine()
                                End If

                                If mode = PS3000AAPI.Mode.AGGREGATED Then
                                    Dim digiValueOR As Short = CShort(&HFF And digiPinnedA(1).Target(i))
                                    digiValueOR <<= 8
                                    digiValueOR = digiValueOR Or CShort(&HFF And digiPinnedA(0).Target(i))
                                    Console.WriteLine("Index={0,4:D}: Bitwise  OR of last {1} values = 0x{2,4:X}  ", i, downsampleRatio, digiValueOR.ToString("X4"))

                                    Dim digiValueAND As Short = CShort(&HFF And digiPinnedB(1).Target(i))

                                    digiValueAND <<= 8
                                    digiValueAND = digiValueAND Or CShort(&HFF And digiPinnedB(0).Target(i))
                                    Console.WriteLine("Index={0,4:D}: Bitwise AND of last {1} values = 0x{2,4:X}  ", i, downsampleRatio, digiValueAND.ToString("X4"))
                                End If
                            Next
                        End If
                    End If
                End If
            End While

            If Console.KeyAvailable Then
                Console.ReadKey(True) '// clear the key
            End If

            PS3000AAPI.StopUnit(_handle)

            If Not IsNothing(writer) Then writer.Close()

            If (Not _autoStop) Then
                Console.WriteLine("data collection aborted")
            End If


            If mode = PS3000AAPI.Mode.ANALOGUE Then
                For Each p As PinnedArray(Of Short) In minPinned
                    If Not IsNothing(p) Then
                        p.Dispose()
                    End If
                Next
                For Each p As PinnedArray(Of Short) In maxPinned
                    If Not IsNothing(p) Then
                        p.Dispose()
                    End If
                Next
            End If

            If mode = PS3000AAPI.Mode.AGGREGATED Then
                For Each p As PinnedArray(Of Short) In digiPinnedA
                    If Not IsNothing(p) Then
                        p.Dispose()
                    End If
                Next
                For Each p As PinnedArray(Of Short) In digiPinnedB
                    If Not IsNothing(p) Then
                        p.Dispose()
                    End If
                Next
            End If

            If mode = PS3000AAPI.Mode.DIGITAL Then
                For Each p As PinnedArray(Of Short) In digiPinnedA
                    If Not IsNothing(p) Then
                        p.Dispose()
                    End If
                Next
            End If
        End Sub

        '/****************************************************************************
        '* CollectStreamingImmediate
        '*  this function demonstrates how to collect a stream of data
        '*  from the unit (start collecting immediately)
        '***************************************************************************/
        Private Sub CollectStreamingImmediate()
            SetDefaults()

            Console.WriteLine("Collect streaming...")
            Console.WriteLine("Data is written to disk file (stream.txt)")
            Console.WriteLine("Press a key to start")
            WaitForKey()

            '/* Trigger disabled	*/
            SetTrigger({}, 0, {}, 0, {}, Nothing, 0, 0, 0)

            StreamDataHandler(0, PS3000AAPI.Mode.ANALOGUE)
        End Sub

        '/****************************************************************************
        '* CollectStreamingTriggered
        '*  this function demonstrates how to collect a stream of data
        '*  from the unit (start collecting on trigger)
        '***************************************************************************/
        Private Sub CollectStreamingTriggered()
            Dim triggerVoltage As Short = mv_to_adc(inputRanges(_channelSettings(PS3000AAPI.Channel.ChannelA).range) / 5, _channelSettings(PS3000AAPI.Channel.ChannelA).range) '// ChannelInfo stores ADC counts

            Dim sourceDetails As PS3000AAPI.TriggerChannelProperties() = New PS3000AAPI.TriggerChannelProperties() {
                New PS3000AAPI.TriggerChannelProperties( _
                    triggerVoltage, _
                    256 * 10, _
                    triggerVoltage, _
                    256 * 10, _
                    PS3000AAPI.Channel.ChannelA, _
                    PS3000AAPI.ThresholdMode.Level)
            }

            Dim conditions As PS3000AAPI.TriggerConditions() = New PS3000AAPI.TriggerConditions() {
                New PS3000AAPI.TriggerConditions( _
                    PS3000AAPI.TriggerState.TrigTrue, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare)
            }

            Dim directions As PS3000AAPI.ThresholdDirection() = New PS3000AAPI.ThresholdDirection() {
                PS3000AAPI.ThresholdDirection.Rising,
                PS3000AAPI.ThresholdDirection.None, _
                PS3000AAPI.ThresholdDirection.None, _
                PS3000AAPI.ThresholdDirection.None, _
                PS3000AAPI.ThresholdDirection.None, _
                PS3000AAPI.ThresholdDirection.None
            }

            Console.WriteLine("Collect streaming triggered...")
            Console.WriteLine("Data is written to disk file (stream.txt)")
            Console.WriteLine("Press a key to start")

            WaitForKey()
            SetDefaults()

            '/* Trigger enabled
            ' * Rising edge
            ' * Threshold = 1/5Range */

            SetTrigger(sourceDetails, 1, conditions, 1, directions, Nothing, 0, 0, 0)
            StreamDataHandler(100000, PS3000AAPI.Mode.ANALOGUE)
        End Sub

        '/****************************************************************************
        '* DisplaySettings 
        '* Displays information about the user configurable settings in this example
        '***************************************************************************/
        Private Sub DisplaySettings()
            Dim ch As Integer
            Dim voltage As Integer

            Console.WriteLine("Readings will be scaled in {0}", IIf(_scaleVoltages, ("mV"), ("ADC counts")))

            For ch = 0 To _channelCount - 1
                If Not _channelSettings(ch).enabled Then
                    Console.WriteLine("Channel {0} Voltage Range = Off", Chr(Asc("A") + ch))
                Else
                    voltage = inputRanges(CInt(_channelSettings(ch).range))
                    Console.Write("Channel {0} Voltage Range = ", Chr(Asc("A") + ch))

                    If (voltage < 1000) Then
                        Console.WriteLine("{0}mV", voltage)
                    Else
                        Console.WriteLine("{0}V", voltage / 1000)
                    End If
                End If
            Next
            Console.WriteLine()
        End Sub

        '/****************************************************************************
        '* DigitalBlockImmediate
        '* Collect a block of data from the digital ports with triggering disabled
        '***************************************************************************/
        Private Sub DigitalBlockImmediate()
            Console.WriteLine("Digital Block Immediate")
            Console.WriteLine("Press a key to start...")
            WaitForKey()

            SetTrigger({}, 0, {}, 0, {}, Nothing, 0, 0, 0, {}, 0)

            BlockDataHandler("First 10 readings", 0, PS3000AAPI.Mode.DIGITAL)
        End Sub

        '/****************************************************************************
        '* DigitalBlockTriggered
        '* Collect a block of data from the digital ports with triggering disabled
        '***************************************************************************/
        Private Sub DigitalBlockTriggered()
            Console.WriteLine("Digital Block Triggered")
            Console.WriteLine("Collect block of data when the trigger occurs...")
            Console.WriteLine("Digital Channel   0   --- Rising")
            Console.WriteLine("Digital Channel   4   --- High")
            Console.WriteLine("Other Digital Channels - Don't Care")

            Console.WriteLine("Press a key to start...")
            WaitForKey()

            Dim conditions() As PS3000AAPI.TriggerConditionsV2 = New PS3000AAPI.TriggerConditionsV2() {
                New PS3000AAPI.TriggerConditionsV2( _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.TrigTrue _
                )}

            Dim digitalDirections(2) As PS3000AAPI.DigitalChannelDirections

            digitalDirections(0).DigiPort = PS3000AAPI.DigitalChannel.PS3000A_DIGITAL_CHANNEL_0
            digitalDirections(0).DigiDirection = PS3000AAPI.DigitalDirection.PS3000A_DIGITAL_DIRECTION_RISING

            digitalDirections(1).DigiPort = PS3000AAPI.DigitalChannel.PS3000A_DIGITAL_CHANNEL_4
            digitalDirections(1).DigiDirection = PS3000AAPI.DigitalDirection.PS3000A_DIGITAL_DIRECTION_HIGH

            SetTrigger({}, 0, conditions, 1, {}, Nothing, 0, 0, 0, digitalDirections, 2)

            BlockDataHandler("First 10 readings", 0, PS3000AAPI.Mode.DIGITAL)
        End Sub

        '/****************************************************************************
        '* ANDAnalogueDigitalTriggered
        '*  this function demonstrates how to combine Digital AND Analogue triggers
        '*  to collect a block of data.
        '****************************************************************************/
        Private Sub ANDAnalogueDigitalTriggered()
            Console.WriteLine("Analogue AND Digital Triggered Block")

            Dim triggerVoltage As Short = mv_to_adc(inputRanges(_channelSettings(PS3000AAPI.Channel.ChannelA).range) / 5, _channelSettings(PS3000AAPI.Channel.ChannelA).range)  ' // ChannelInfo stores ADC counts

            Dim sourceDetails() As PS3000AAPI.TriggerChannelProperties = New PS3000AAPI.TriggerChannelProperties() {
                New PS3000AAPI.TriggerChannelProperties(triggerVoltage, _
                    256 * 10, _
                    triggerVoltage, _
                    256 * 10, _
                    PS3000AAPI.Channel.ChannelA, _
                    PS3000AAPI.ThresholdMode.Level)
            }


            Dim conditions() As PS3000AAPI.TriggerConditionsV2 = New PS3000AAPI.TriggerConditionsV2() {
                New PS3000AAPI.TriggerConditionsV2( _
                    PS3000AAPI.TriggerState.TrigTrue, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.DontCare, _
                    PS3000AAPI.TriggerState.TrigTrue)
            }

            Dim directions() As PS3000AAPI.ThresholdDirection = New PS3000AAPI.ThresholdDirection() {
                PS3000AAPI.ThresholdDirection.Above,
                PS3000AAPI.ThresholdDirection.None,
                PS3000AAPI.ThresholdDirection.None,
                PS3000AAPI.ThresholdDirection.None,
                PS3000AAPI.ThresholdDirection.None,
                PS3000AAPI.ThresholdDirection.None
            }


            Dim digitalDirections(2) As PS3000AAPI.DigitalChannelDirections

            digitalDirections(0).DigiPort = PS3000AAPI.DigitalChannel.PS3000A_DIGITAL_CHANNEL_0
            digitalDirections(0).DigiDirection = PS3000AAPI.DigitalDirection.PS3000A_DIGITAL_DIRECTION_RISING

            digitalDirections(1).DigiPort = PS3000AAPI.DigitalChannel.PS3000A_DIGITAL_CHANNEL_4
            digitalDirections(1).DigiDirection = PS3000AAPI.DigitalDirection.PS3000A_DIGITAL_DIRECTION_HIGH

            Console.Write("Collect a block of data when value is above {0}", _
                IIf(_scaleVoltages, _
                adc_to_mv(sourceDetails(0).ThresholdMajor, _channelSettings(PS3000AAPI.Channel.ChannelA).range), _
                sourceDetails(0).ThresholdMajor) _
            )
            Console.WriteLine("{0}", IIf(_scaleVoltages, "mV ", "ADC Counts "))
            Console.WriteLine("AND ")
            Console.WriteLine("Digital Channel  0   --- Rising")
            Console.WriteLine("Digital Channel  4   --- High")
            Console.WriteLine("Other Digital Channels - Don't Care")
            Console.WriteLine()
            Console.WriteLine("Press a key to start...")
            WaitForKey()

            SetDefaults()

            '/* Trigger enabled
            '* Channel A
            '* Rising edge
            '* Threshold = 1000mV 
            '* Digial
            '* Ch0 Rising
            '* Ch4 High */

            SetTrigger(sourceDetails, 1, conditions, 1, directions, Nothing, 0, 0, 0, digitalDirections, 2)

            BlockDataHandler("Ten readings after trigger", 0, PS3000AAPI.Mode.MIXED)

            DisableAnalogue()
        End Sub

        '/****************************************************************************
        '* ORAnalogueDigitalTriggered
        '*  this function demonstrates how to trigger off Digital OR Analogue signals
        '*  to collect a block of data.
        '****************************************************************************/
        Private Sub ORAnalogueDigitalTriggered()
            Console.WriteLine("Analogue OR Digital Triggered Block")
            Console.WriteLine("Collect block of data when an Analogue OR Digital triggers occurs...")

            Dim triggerVoltage As Short = mv_to_adc(inputRanges(_channelSettings(PS3000AAPI.Channel.ChannelA).range) / 5, _channelSettings(PS3000AAPI.Channel.ChannelA).range) '// ChannelInfo stores ADC counts
            Dim sourceDetails() As PS3000AAPI.TriggerChannelProperties = New PS3000AAPI.TriggerChannelProperties() {
                New PS3000AAPI.TriggerChannelProperties(triggerVoltage,
                                             256 * 10,
                                             triggerVoltage,
                                             256 * 10,
                                             PS3000AAPI.Channel.ChannelA,
                                             PS3000AAPI.ThresholdMode.Level)}


            Dim conditions(2) As PS3000AAPI.TriggerConditionsV2

            conditions(0).ChannelA = PS3000AAPI.TriggerState.TrigTrue
            conditions(0).ChannelB = PS3000AAPI.TriggerState.DontCare
            conditions(0).ChannelC = PS3000AAPI.TriggerState.DontCare
            conditions(0).ChannelD = PS3000AAPI.TriggerState.DontCare
            conditions(0).External = PS3000AAPI.TriggerState.DontCare
            conditions(0).Aux = PS3000AAPI.TriggerState.DontCare
            conditions(0).Pwq = PS3000AAPI.TriggerState.DontCare
            conditions(0).Digital = PS3000AAPI.TriggerState.DontCare

            conditions(1).ChannelA = PS3000AAPI.TriggerState.DontCare
            conditions(1).ChannelB = PS3000AAPI.TriggerState.DontCare
            conditions(1).ChannelC = PS3000AAPI.TriggerState.DontCare
            conditions(1).ChannelD = PS3000AAPI.TriggerState.DontCare
            conditions(1).External = PS3000AAPI.TriggerState.DontCare
            conditions(1).Aux = PS3000AAPI.TriggerState.DontCare
            conditions(1).Pwq = PS3000AAPI.TriggerState.DontCare
            conditions(1).Digital = PS3000AAPI.TriggerState.TrigTrue

            Dim directions() As PS3000AAPI.ThresholdDirection
            directions = {
                PS3000AAPI.ThresholdDirection.Rising,
                PS3000AAPI.ThresholdDirection.None,
                PS3000AAPI.ThresholdDirection.None,
                PS3000AAPI.ThresholdDirection.None,
                PS3000AAPI.ThresholdDirection.None,
                PS3000AAPI.ThresholdDirection.None
            }

            Dim digitalDirections(2) As PS3000AAPI.DigitalChannelDirections

            digitalDirections(0).DigiPort = PS3000AAPI.DigitalChannel.PS3000A_DIGITAL_CHANNEL_0
            digitalDirections(0).DigiDirection = PS3000AAPI.DigitalDirection.PS3000A_DIGITAL_DIRECTION_RISING

            digitalDirections(1).DigiPort = PS3000AAPI.DigitalChannel.PS3000A_DIGITAL_CHANNEL_4
            digitalDirections(1).DigiDirection = PS3000AAPI.DigitalDirection.PS3000A_DIGITAL_DIRECTION_HIGH


            Console.Write( _
                "Collect a block of data when value rises past {0}", _
                IIf(
                    _scaleVoltages,
                    adc_to_mv(
                        sourceDetails(0).ThresholdMajor,
                        _channelSettings(CInt(PS3000AAPI.Channel.ChannelA)).range
                    ),
                    sourceDetails(0).ThresholdMajor
                )
            )

            Console.WriteLine("{0}", IIf(_scaleVoltages, "mV ", "ADC Counts "))
            Console.WriteLine("OR")
            Console.WriteLine("Digital Channel  0   --- Rising")
            Console.WriteLine("Digital Channel  4   --- High")
            Console.WriteLine("Other Digital Channels - Don't Care")
            Console.WriteLine()
            Console.WriteLine("Press a key to start...")
            WaitForKey()

            SetDefaults()

            '/* Trigger enabled
            ' * Channel A
            ' * Rising edge
            ' * Threshold = 1000mV 
            ' * Digial
            ' * Ch0 Rising
            ' * Ch4 High */
            SetTrigger(sourceDetails, 1, conditions, 2, directions, Nothing, 0, 0, 0, digitalDirections, 2)
            BlockDataHandler("Ten readings after trigger", 0, PS3000AAPI.Mode.MIXED)
            DisableAnalogue()
        End Sub

        '/****************************************************************************
        '* DigitalStreamingImmediate
        '* Streams data from the digital ports with triggering disabled
        '***************************************************************************/
        Private Sub DigitalStreamingImmediate()
            Console.WriteLine("Digital Streaming Immediate....")
            Console.WriteLine("Press a key to start")
            WaitForKey()

            SetTrigger({}, 0, {}, 0, {}, Nothing, 0, 0, 0, {}, 0)
            StreamDataHandler(0, PS3000AAPI.Mode.DIGITAL)
        End Sub

        '/****************************************************************************
        '* DigitalStreamingImmediate
        '* Streams data from the digital ports with triggering disabled
        '***************************************************************************/
        Private Sub DigitalStreamingAggregated()
            Console.WriteLine("Digital Streaming with Aggregation....")
            Console.WriteLine("Press a key to start")
            WaitForKey()

            SetTrigger({}, 0, {}, 0, {}, Nothing, 0, 0, 0, {}, 0)
            StreamDataHandler(0, PS3000AAPI.Mode.AGGREGATED)
        End Sub

        '/****************************************************************************
        ' * DigitalMenu - Only shown for the MSO scopes
        ' ****************************************************************************/
        Public Sub DigitalMenu()
            Dim ch As Char

            DisableAnalogue()                  '// Disable the analogue ports
            SetDigitals()                      '// Enable digital ports & set logic level

            ch = " "
            While (ch <> "X")
                Console.WriteLine()
                Console.WriteLine("Digital Port Menu")
                Console.WriteLine()
                Console.WriteLine("B - Digital Block Immediate")
                Console.WriteLine("T - Digital Block Triggered")
                Console.WriteLine("A - Analogue 'AND' Digital Triggered Block")
                Console.WriteLine("O - Analogue 'OR'  Digital Triggered Block")
                Console.WriteLine("S - Digital Streaming Mode")
                Console.WriteLine("V - Digital Streaming Aggregated")
                Console.WriteLine("X - Return to previous menu")
                Console.WriteLine()
                Console.WriteLine("Operation:")

                ch = Char.ToUpper(Console.ReadKey(True).KeyChar)

                Console.WriteLine()
                Select Case (ch)
                    Case "B"
                        DigitalBlockImmediate()
                    Case "T"
                        DigitalBlockTriggered()
                    Case "A"
                        ANDAnalogueDigitalTriggered()
                    Case "O"
                        ORAnalogueDigitalTriggered()
                    Case "S"
                        DigitalStreamingImmediate()
                    Case "V"
                        DigitalStreamingAggregated()
                End Select
            End While

            DisableDigital()       '// disable the digital ports when we're finished
        End Sub

        '/****************************************************************************
        '* Run - show menu and call user selected options
        '****************************************************************************/
        Public Sub Run()
            '// setup devices
            GetDeviceInfo()
            _timebase = 1

            _firstRange = PS3000AAPI.Range.Range_100MV
            _lastRange = PS3000AAPI.Range.Range_20V
            ReDim _channelSettings(MAX_CHANNELS)

            Dim i As Integer
            For i = 0 To MAX_CHANNELS - 1
                _channelSettings(i).enabled = False
                _channelSettings(i).DCcoupled = True
                _channelSettings(i).range = PS3000AAPI.Range.Range_5V
            Next
            _channelSettings(0).enabled = True

            '// main loop - read key and call routine
            Dim ch As Char = " "
            While (ch <> "X")
                DisplaySettings()

                Console.WriteLine("")
                Console.WriteLine("B - Immediate block              V - Set voltages")
                Console.WriteLine("T - Triggered block              I - Set timebase")
                Console.WriteLine("R - Rapid block                  A - ADC counts/mV")
                Console.WriteLine("S - Immediate Streaming")
                Console.WriteLine("W - Triggered Streaming")
                Console.WriteLine(IIf(_digitalPorts > 0, "D - Digital Ports Menu", ""))
                Console.WriteLine("                                 X - exit")
                Console.WriteLine("Operation:")

                ch = Char.ToUpper(Console.ReadKey(True).KeyChar)

                Console.WriteLine()

                Select Case ch
                    Case "B"
                        CollectBlockImmediate()

                    Case "T"
                        CollectBlockTriggered()

                    Case "R"
                        CollectBlockRapid()

                    Case "S"
                        CollectStreamingImmediate()

                    Case "W"
                        CollectStreamingTriggered()

                    Case "V"
                        SetVoltages()

                    Case "I"
                        SetTimebase()

                    Case "A"
                        _scaleVoltages = Not _scaleVoltages

                    Case "D"
                        If _digitalPorts > 0 Then
                            DigitalMenu()
                        End If

                    Case "X"
                        '/* Handled by outer loop */

                    Case Else
                        Console.WriteLine("Invalid operation")
                End Select
            End While
        End Sub

        Public Sub New(handle As Short)
            _handle = handle
        End Sub

    End Class

    Module Module1
        '/****************************************************************************
        '*  WaitForKey
        '*  Wait for user's keypress
        '****************************************************************************/
        Public Sub WaitForKey()
            While (Not Console.KeyAvailable)
                Thread.Sleep(100)
            End While

            If Console.KeyAvailable Then
                Console.ReadKey(True) '// clear the key 
            End If
        End Sub

        '/****************************************************************************
        '* PowerSourceSwitch - Handle status errors connected with the power source
        '****************************************************************************/
        Public Function PowerSourceSwitch(handle As Short, status As Short) As Short
            Dim ch As Char

            Select Case status
                Case PS3000AAPI.PICO_POWER_SUPPLY_NOT_CONNECTED
                    Do
                        Console.WriteLine(String.Join("5V Power Supply not connected", "Do you want to run using power from the USB lead?", vbCrLf))
                        ch = Char.ToUpper(Console.ReadKey(True).KeyChar)
                        If ch = "Y" Then
                            Console.WriteLine("Powering the uint via USB")
                            status = PS3000AAPI.ChangePowerSource(handle, status)
                        End If
                    Loop While (ch <> "Y" And ch <> "N")

                    Console.WriteLine(IIf(ch = "N", "Please use the 5V power supply to power this unit", ""))

                Case PS3000AAPI.PICO_POWER_SUPPLY_CONNECTED
                    Console.WriteLine("Using 5V power supply voltage")
                    status = PS3000AAPI.ChangePowerSource(handle, status)

                Case PS3000AAPI.PICO_POWER_SUPPLY_UNDERVOLTAGE
                    Do
                        Console.WriteLine("")
                        Console.WriteLine("USB not supplying required voltage")
                        Console.WriteLine("Please plug in the +5V power supply,")
                        Console.WriteLine("then hit a key to continue, or Esc to exit...")
                        ch = Console.ReadKey().KeyChar
                        If ch = Chr(&H1B) Then
                            Environment.Exit(0)
                        End If
                        status = PowerSourceSwitch(handle, PS3000AAPI.PICO_POWER_SUPPLY_CONNECTED)
                    Loop While (status = PS3000AAPI.PICO_POWER_SUPPLY_REQUEST_INVALID)
            End Select
            Return status
        End Function

        Private Function deviceOpen(ByRef handle As Short) As Short
            Dim status As Short = PS3000AAPI.OpenUnit(handle, Nothing)

            If status <> PS3000AAPI.PICO_OK Then
                status = PowerSourceSwitch(handle, status)

                If status = PS3000AAPI.PICO_POWER_SUPPLY_UNDERVOLTAGE Then
                    status = PowerSourceSwitch(handle, status)
                End If
            End If

            If status <> PS3000AAPI.PICO_OK Then
                Console.WriteLine("Unable to open device")
                Console.WriteLine("Error code : 0x{0}", Convert.ToString(status, 16))
                WaitForKey()
            Else
                Console.WriteLine("Handle: {0}", handle)
            End If

            Return status
        End Function

        Public Sub Main()
            Console.WriteLine("VB.Net PS3000A driver example program")
            Console.WriteLine("Version 1.0")
            Dim handle As Short

            Console.WriteLine("Opening the device...")
            If deviceOpen(handle) = 0 Then
                Console.WriteLine("Device opened successfully")

                Dim consoleExample As PS3000AConsole = New PS3000AConsole(handle)
                consoleExample.Run()

                PS3000AAPI.CloseUnit(handle)
            End If
        End Sub

    End Module

End Namespace
