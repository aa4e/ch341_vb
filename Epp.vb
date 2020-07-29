Imports System.Runtime.InteropServices

Namespace Ch341

    ''' <summary>
    ''' Работа с CH341 в параллельном режиме EPP.
    ''' </summary>
    ''' <remarks>
    ''' Pins
    ''' 15..22  D0..D7  Bi-dir
    ''' 25      WRITE#    Out
    ''' 4       DATAS#    Out
    ''' 26      RST#      Out
    ''' 3       ADDRS#    Out
    ''' 27      WAIT#     In
    ''' 7       INT#      In
    ''' 5       ERR#/IN0  In
    ''' 8       SLCT/IN1  In 
    ''' 9       PEMP/IN2  In
    ''' </remarks>
    Public NotInheritable Class Epp
        Inherits Parallel

#Region "CTOR"

        ''' <summary>
        ''' Открывает устройство в режиме EPP.
        ''' </summary>
        ''' <param name="index"></param>
        ''' <param name="isExclusive"></param>
        ''' <param name="mode">Режим, кроме MEM.</param>
        Public Sub New(index As Integer, Optional isExclusive As Boolean = False, Optional mode As ParallelMode = ParallelMode.EPP17)
            MyBase.New(index, isExclusive)
            If (mode = ParallelMode.MEM) Then
                mode = ParallelMode.EPP17
            End If
            InitParallel(mode)
        End Sub

#End Region '/CTOR

#Region "ENUM"

        ''' <summary>
        ''' Регистры параллельного порта.
        ''' </summary>
        Public Enum Registers As Integer
            Data
            Address
        End Enum

#End Region '/ENUM

#Region "EPP"

        ''' <remarks>EPP mode read data: WR#=1, DS#=0, AS#=1, D0-D7=input.</remarks>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341EppReadData(iIndex As Integer, oBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <remarks>EPP mode read address: WR#=1, DS#=1, AS#=0, D0-D7=input.</remarks>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341EppReadAddr(iIndex As Integer, oBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Прочитать регистр данных или адреса EPP.
        ''' </summary>
        ''' <param name="lenToRead">Число байтов для чтения.</param>
        ''' <param name="register">Регистр, из которого нужно прочитать данные.</param>
        Public Function EppReadRegister(lenToRead As Integer, Optional register As Registers = Registers.Data) As Byte()
            Dim readBuffer(lenToRead - 1) As Byte
            Dim res As Boolean = False
            Select Case register
                Case Registers.Data
                    res = CH341EppReadData(DeviceIndex, readBuffer, lenToRead)
                Case Registers.Address
                    res = CH341EppReadAddr(DeviceIndex, readBuffer, lenToRead)
            End Select
            If res Then
                Return readBuffer
            End If
            Throw New Exception("Ошибка чтения EPP.")
        End Function

        ''' <remarks>EPP way to write data: WR#=0, DS#=0, AS#=1, D0-D7=output.</remarks>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341EppWriteData(iIndex As Integer, iBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <remarks>EPP mode write address: WR#=0, DS#=1, AS#=0, D0-D7=output.</remarks>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341EppWriteAddr(iIndex As Integer, iBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Записыват в регистр данных или адреса массив байтов и возвращает число записанных байтов.
        ''' </summary>
        ''' <param name="dataToWrite">Данные для записи.</param>
        ''' <param name="register">Регистр, в который записываются данные.</param>
        Public Function EppWriteRegister(dataToWrite As Byte(), Optional register As Registers = Registers.Data) As Integer
            Dim wroteLen As Integer = dataToWrite.Length
            Dim res As Boolean = False
            Select Case register
                Case Registers.Data
                    res = CH341EppWriteData(DeviceIndex, dataToWrite, wroteLen)
                Case Registers.Address
                    res = CH341EppWriteAddr(DeviceIndex, dataToWrite, wroteLen)
            End Select
            If res Then
                Return wroteLen
            End If
            Throw New Exception("Ошибка записи EPP.")
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341EppSetAddr(iIndex As Integer, ByRef iAddr As Byte) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Задаёт адрес EPP.
        ''' </summary>
        ''' <param name="address">Адрес.</param>
        ''' <remarks>WR#=0, DS#=1, AS#=0, D0-D7=output.</remarks>
        Public Sub EppSetAddress(address As Byte)
            Dim res As Boolean = CH341EppSetAddr(DeviceIndex, address)
            If res Then
                Return
            End If
            Throw New Exception("Ошибка выставления адреса EPP.")
        End Sub

#End Region '/EPP

    End Class

End Namespace