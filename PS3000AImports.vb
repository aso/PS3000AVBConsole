'/**************************************************************************
'*
'* Filename:    PS3000AImports.cs
'*
'* Copyright:   Pico Technology Limited 2011
'*
'* Author:      MJL
'*
'* Description:
'*   This file contains all the .NET wrapper calls needed to support
'*   the console example. It also has the enums and structs required
'*   by the (wrapped) function calls.
'*
'* History:
'*    14Dec06	MJL	Created
'* 	 15Oct09	RPM Modified for PS4000
'* 	 23Nov11    CPY Modified for PS3000A

'*
'* Revision Info: "file %n date %f revision %v"
'*						""
'*
'***************************************************************************/

Imports System
Imports System.Runtime.InteropServices
Imports System.Text

Namespace picotech
    Class PS3000AAPI

#Region "Constants"
        Private Const _DRIVER_FILENAME As String = "ps3000A.dll"

        Public Const MaxValue As Int16 = -32512          ' &H7F00
        Public Const MinValue As Int16 = 32512         ' &H8100
        Public Const MaxLogicLevel As Int16 = 32767     ' &H7FFF
        Public Const MinLogicLevel As Int16 = -32767    ' &H8001

        Public Const PICO_OK As UInt32 = 0
        Public Const PICO_CANCELLED As UInt32 = &H3A
        Public Const PICO_POWER_SUPPLY_CONNECTED As UInt32 = &H119
        Public Const PICO_POWER_SUPPLY_NOT_CONNECTED As UInt32 = &H11A
        Public Const PICO_POWER_SUPPLY_REQUEST_INVALID As UInt32 = &H11B
        Public Const PICO_POWER_SUPPLY_UNDERVOLTAGE As UInt32 = &H11C
#End Region

#Region "Driver enums"
        Public Enum Channel As Short
            ChannelA
            ChannelB
            ChannelC
            ChannelD
            External
            Aux
            None
            PS3000A_DIGITAL_PORT0 = &H80    '// digital channel 0 - 7
            PS3000A_DIGITAL_PORT1           '// digital channel 8 - 15
            PS3000A_DIGITAL_PORT2           '// digital channel 16 - 23
            PS3000A_DIGITAL_PORT3           '// digital channel 24 - 31
        End Enum

        Public Enum Range As Short
            Range_10MV
            Range_20MV
            Range_50MV
            Range_100MV
            Range_200MV
            Range_500MV
            Range_1V
            Range_2V
            Range_5V
            Range_10V
            Range_20V
            Range_50V
            Range_MAX_RANGE
        End Enum

        Public Enum ReportedTimeUnits As Short
            FemtoSeconds
            PicoSeconds
            NanoSeconds
            MicroSeconds
            MilliSeconds
            Seconds
        End Enum

        Public Enum ThresholdMode As Short
            Level
            Window
        End Enum

        Public Enum ThresholdDirection As Short
            '// Values for level threshold mode
            '//
            Above
            Below
            Rising
            Falling
            RisingOrFalling

            '// Values for window threshold mode
            '//
            Inside = Above
            Outside = Below
            Enter = Rising
            ThExit = Falling        ' Original "Exit"
            EnterOrExit = RisingOrFalling
            PositiveRunt = 9
            NegativeRunt

            None = Rising
        End Enum

        Public Enum PulseWidthType As Short
            None
            LessThan
            GreaterThan
            InRange
            OutOfRange
        End Enum

        Public Enum TriggerState As Short
            DontCare
            TrigTrue
            TrigFalse
        End Enum

        Public Enum RatioMode As Short
            None
            Aggregate = 1
            Decimate = 2
            Average = 4
        End Enum

        Public Enum DigitalDirection As Short
            PS3000A_DIGITAL_DONT_CARE
            PS3000A_DIGITAL_DIRECTION_LOW
            PS3000A_DIGITAL_DIRECTION_HIGH
            PS3000A_DIGITAL_DIRECTION_RISING
            PS3000A_DIGITAL_DIRECTION_FALLING
            PS3000A_DIGITAL_DIRECTION_RISING_OR_FALLING
            PS3000A_DIGITAL_MAX_DIRECTION
        End Enum

        Public Enum Mode As Short
            ANALOGUE
            DIGITAL
            AGGREGATED
            MIXED
        End Enum

        Public Enum DigitalChannel As Short
            PS3000A_DIGITAL_CHANNEL_0
            PS3000A_DIGITAL_CHANNEL_1
            PS3000A_DIGITAL_CHANNEL_2
            PS3000A_DIGITAL_CHANNEL_3
            PS3000A_DIGITAL_CHANNEL_4
            PS3000A_DIGITAL_CHANNEL_5
            PS3000A_DIGITAL_CHANNEL_6
            PS3000A_DIGITAL_CHANNEL_7
            PS3000A_DIGITAL_CHANNEL_8
            PS3000A_DIGITAL_CHANNEL_9
            PS3000A_DIGITAL_CHANNEL_10
            PS3000A_DIGITAL_CHANNEL_11
            PS3000A_DIGITAL_CHANNEL_12
            PS3000A_DIGITAL_CHANNEL_13
            PS3000A_DIGITAL_CHANNEL_14
            PS3000A_DIGITAL_CHANNEL_15
            PS3000A_DIGITAL_CHANNEL_16
            PS3000A_DIGITAL_CHANNEL_17
            PS3000A_DIGITAL_CHANNEL_18
            PS3000A_DIGITAL_CHANNEL_19
            PS3000A_DIGITAL_CHANNEL_20
            PS3000A_DIGITAL_CHANNEL_21
            PS3000A_DIGITAL_CHANNEL_22
            PS3000A_DIGITAL_CHANNEL_23
            PS3000A_DIGITAL_CHANNEL_24
            PS3000A_DIGITAL_CHANNEL_25
            PS3000A_DIGITAL_CHANNEL_26
            PS3000A_DIGITAL_CHANNEL_27
            PS3000A_DIGITAL_CHANNEL_28
            PS3000A_DIGITAL_CHANNEL_29
            PS3000A_DIGITAL_CHANNEL_30
            PS3000A_DIGITAL_CHANNEL_31
        End Enum

#End Region

#Region "structs"

        <StructLayout(LayoutKind.Sequential, Pack:=1)> _
        Public Structure TriggerChannelProperties
            Public ThresholdMajor As Short
            Public HysteresisMajor As UShort
            Public ThresholdMinor As Short
            Public HysteresisMinor As UShort
            Public Channel As Channel
            Public ThresholdMode As ThresholdMode

            Public Sub New( _
                    ByVal thresholdMajor As Short, _
                    ByVal hysteresisMajor As UShort, _
                    ByVal thresholdMinor As Short, _
                    ByVal hysteresisMinor As UShort, _
                    ByVal channel As Channel, _
                    ByVal thresholdMode As ThresholdMode)
                Me.ThresholdMajor = thresholdMajor
                Me.HysteresisMajor = hysteresisMajor
                Me.ThresholdMinor = thresholdMinor
                Me.HysteresisMinor = hysteresisMinor
                Me.Channel = channel
                Me.ThresholdMode = thresholdMode
            End Sub
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=1)> _
        Public Structure TriggerConditions
            Public ChannelA As TriggerState
            Public ChannelB As TriggerState
            Public ChannelC As TriggerState
            Public ChannelD As TriggerState
            Public External As TriggerState
            Public Aux As TriggerState
            Public Pwq As TriggerState

            Public Sub New( _
              channelA As TriggerState, _
              channelB As TriggerState, _
              channelC As TriggerState, _
              channelD As TriggerState, _
              external As TriggerState, _
              aux As TriggerState, _
              pwq As TriggerState)
                Me.ChannelA = channelA
                Me.ChannelB = channelB
                Me.ChannelC = channelC
                Me.ChannelD = channelD
                Me.External = external
                Me.Aux = aux
                Me.Pwq = pwq
            End Sub
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=1)> _
        Public Structure TriggerConditionsV2
            Public ChannelA As TriggerState
            Public ChannelB As TriggerState
            Public ChannelC As TriggerState
            Public ChannelD As TriggerState
            Public External As TriggerState
            Public Aux As TriggerState
            Public Pwq As TriggerState
            Public Digital As TriggerState

            Public Sub New( _
                    channelA As TriggerState, _
                    channelB As TriggerState, _
                    channelC As TriggerState, _
                    channelD As TriggerState, _
                    external As TriggerState, _
                    aux As TriggerState, _
                    pwq As TriggerState, _
                    digital As TriggerState)
                Me.ChannelA = channelA
                Me.ChannelB = channelB
                Me.ChannelC = channelC
                Me.ChannelD = channelD
                Me.External = external
                Me.Aux = aux
                Me.Pwq = pwq
                Me.Digital = digital
            End Sub
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=1)> _
        Public Structure PwqConditions
            Public ChannelA As TriggerState
            Public ChannelB As TriggerState
            Public ChannelC As TriggerState
            Public ChannelD As TriggerState
            Public External As TriggerState
            Public Aux As TriggerState

            Public Sub New( _
                    channelA As TriggerState, _
                    channelB As TriggerState, _
                    channelC As TriggerState, _
                    channelD As TriggerState, _
                    external As TriggerState, _
                    aux As TriggerState)
                Me.ChannelA = channelA
                Me.ChannelB = channelB
                Me.ChannelC = channelC
                Me.ChannelD = channelD
                Me.External = external
                Me.Aux = aux
            End Sub
        End Structure

        <StructLayout(LayoutKind.Sequential, Pack:=1)> _
        Public Structure DigitalChannelDirections
            Public DigiPort As DigitalChannel
            Public DigiDirection As DigitalDirection

            Public Sub New( _
                 digiPort As DigitalChannel, _
                 digiDirection As DigitalDirection)

                Me.DigiPort = digiPort
                Me.DigiDirection = digiDirection
            End Sub
        End Structure
#End Region

#Region "Callback delegates"
        Public Delegate Sub ps3000aBlockReady( _
            handle As Short, _
            status As Short, _
            pVoid As IntPtr)

        Public Delegate Sub ps3000aStreamingReady( _
            handle As Short, _
            noOfSamples As Int32, _
            startIndex As UInt32, _
            ov As Short, _
            triggerAt As UInt32, _
            triggered As Short, _
            autoStop As Short, _
            pVoid As IntPtr)

        Public Delegate Sub ps3000DataReady( _
            handle As Short, _
            status As Short, _
            noOfSamples As Int32, _
            overflow As Short, _
            pVoid As IntPtr)

#End Region

#Region "Driver Imports"
        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aOpenUnit")>
        Public Shared Function OpenUnit( _
                ByRef handle As Short, _
                serial As StringBuilder) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aCloseUnit")>
        Public Shared Function CloseUnit( _
            ByVal handle As Short) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aRunBlock")>
        Public Shared Function RunBlock( _
                ByVal handle As Short, _
                noOfPreTriggerSamples As Int32, _
                noOfPostTriggerSamples As Int32, _
                timebase As UInt32, _
                oversample As Short, _
                ByRef timeIndisposedMs As Int32, _
                segmentIndex As UShort, _
                lpps3000aBlockReady As ps3000aBlockReady, _
                pVoid As IntPtr) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aStop")>
        Public Shared Function StopUnit( _
                handle As Short) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aSetChannel")>
        Public Shared Function SetChannel( _
                handle As Short, _
                channel As Channel, _
                enabled As Short, _
                dc As Short, _
                range As Range, _
                analogueOffset As Single) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aSetDataBuffer")>
        Public Shared Function SetDataBuffer( _
                handle As Short, _
                channel As Channel, _
                buffer As Short(), _
                bufferLth As Int32, _
                segmentIndex As UShort, _
                ratioMode As RatioMode) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aSetDataBuffers")>
        Public Shared Function SetDataBuffers( _
                handle As Short, _
                channel As Channel, _
                bufferMax As Short(), _
                bufferMin As Short(), _
                bufferLth As Int32, _
                segmentIndex As UShort, _
                ratioMode As RatioMode) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aSetTriggerChannelDirections")>
        Public Shared Function SetTriggerChannelDirections(
                handle As Short, _
                channelA As ThresholdDirection, _
                channelB As ThresholdDirection, _
                channelC As ThresholdDirection, _
                channelD As ThresholdDirection, _
                ext As ThresholdDirection, _
                aux As ThresholdDirection) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aGetTimebase")>
        Public Shared Function GetTimebase( _
                handle As Short, _
                timebase As Int32, _
                noSamples As Int32, _
                ByRef timeIntervalNanoseconds As Int32, _
                oversample As Short, _
                ByRef maxSamples As Int32, _
                segmentIndex As UShort) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aGetValues")>
        Public Shared Function GetValues( _
                handle As Short, _
                startIndex As UInt32, _
                ByRef noOfSamples As UInt32, _
                downSampleRatio As UInt32, _
                downSampleRatioMode As RatioMode, _
                segmentIndex As UShort, _
                ByRef overflow As Short) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aSetPulseWidthQualifier")>
        Public Shared Function SetPulseWidthQualifier( _
                handle As Short, _
                conditions() As PwqConditions, _
                nConditions As Short, _
                direction As ThresholdDirection, _
                lower As UInt32, _
                upper As UInt32, _
                type As PulseWidthType) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aSetTriggerChannelProperties")>
        Public Shared Function SetTriggerChannelProperties( _
                handle As Short, _
                channelProperties() As TriggerChannelProperties, _
                nChannelProperties As Short, _
                auxOutputEnable As Short, _
                autoTriggerMilliseconds As Int32) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aSetTriggerChannelConditions")>
        Public Shared Function SetTriggerChannelConditions( _
                handle As Short, _
                conditions() As TriggerConditions, _
                nConditions As Short) As Short
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aSetTriggerChannelConditionsV2")>
        Public Shared Function SetTriggerChannelConditionsV2( _
                handle As Short, _
                conditions() As TriggerConditionsV2, _
                nConditions As Short) As Short
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aSetTriggerDelay")>
        Public Shared Function SetTriggerDelay( _
                handle As Short, _
                delay As UInt32) As Short
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aGetUnitInfo")>
        Public Shared Function GetUnitInfo( _
                handle As Short, _
                infoString As StringBuilder, _
                stringLength As Short, _
                ByRef requiredSize As Short, _
                info As Int32) As Short
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aRunStreaming")>
        Public Shared Function RunStreaming( _
                handle As Short, _
                ByRef sampleInterval As UInt32, _
                sampleIntervalTimeUnits As ReportedTimeUnits, _
                maxPreTriggerSamples As UInt32, _
                maxPostPreTriggerSamples As UInt32, _
                autoStop As Boolean, _
                downSamplingRatio As UInt32, _
                downSampleRatioMode As RatioMode, _
                overviewBufferSize As UInt32) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aGetStreamingLatestValues")>
        Public Shared Function GetStreamingLatestValues( _
                handle As Short, _
                lpps3000aStreamingReady As ps3000aStreamingReady, _
                pVoid As IntPtr) As Short
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aSetNoOfCaptures")>
        Public Shared Function SetNoOfRapidCaptures( _
                handle As Short, _
                nCaptures As UShort) As Short
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aMemorySegments")>
        Public Shared Function MemorySegments( _
                handle As Short, _
                nSegments As UShort, _
                ByRef nMaxSamples As Int32) As Short
        End Function


        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aGetValuesBulk")>
        Public Shared Function GetValuesRapid( _
                handle As Short, _
                ByRef noOfSamples As UInt32, _
                fromSegmentIndex As UShort, _
                toSegmentIndex As UShort, _
                downSampleRatio As UInt32, _
                downSampleRatioMode As RatioMode, _
                overflow As Short()) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aChangePowerSource")>
        Public Shared Function ChangePowerSource(
                handle As Short, _
                status As Short) As Short
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aSetDigitalPort")>
        Public Shared Function SetDigitalPort( _
                handle As Short, _
                digiPort As Channel, _
                enabled As Short, _
                logicLevel As Short) As Short
            '
        End Function

        <DllImport(_DRIVER_FILENAME, EntryPoint:="ps3000aSetTriggerDigitalPortProperties")>
        Public Shared Function SetTriggerDigitalPort( _
                handle As Short, _
                DigiChannelDirections() As DigitalChannelDirections, _
                nDigiChannelDirections As Short) As Short
            '
        End Function

#End Region
    End Class
End Namespace
