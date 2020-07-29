Imports System.Runtime.InteropServices

Namespace Ch341

    ''' <summary>
    ''' Работа с CH341 в параллельном режиме MEM.
    ''' </summary>
    Public NotInheritable Class Mem
        Inherits Parallel

#Region "CTOR"

        ''' <summary>
        ''' Открывает устройство в параллельном режиме MEM.
        ''' </summary>
        ''' <param name="index"></param>
        ''' <param name="isExclusive"></param>
        Public Sub New(index As Integer, Optional isExclusive As Boolean = False)
            MyBase.New(index, isExclusive)
            InitParallel(ParallelMode.MEM)
        End Sub

#End Region '/CTOR

#Region "MEM"

        ''' <remarks>MEM mode read address 0: WR#=1, DS#/RD#=0, AS#/ADDR=0, D0-D7=input.</remarks>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341MemReadAddr0(ByVal iIndex As Integer, ByVal oBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <remarks>MEM mode read address 1: WR#=1, DS#/RD#=0, AS#/ADDR=1, D0-D7=input.</remarks>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341MemReadAddr1(ByVal iIndex As Integer, ByVal oBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Читает заданное число байтов из одного из двух регистров.
        ''' </summary>
        ''' <param name="lenToRead">Сколько байтов прочитать.</param>
        ''' <param name="addr0">Читаем из адреса A0 (true) или A1 (false). Выбор адреса позволяет читать из одного из двух регистров.</param>
        Public Function MemRead(ByVal lenToRead As Integer, Optional ByVal addr0 As Boolean = True) As Byte()
            Dim res As Boolean = False
            Dim readBuffer(lenToRead - 1) As Byte
            If addr0 Then
                res = CH341MemReadAddr0(DeviceIndex, readBuffer, lenToRead)
            Else
                res = CH341MemReadAddr1(DeviceIndex, readBuffer, lenToRead)
            End If
            If res Then
                Return readBuffer
            End If
            Throw New Exception("Ошибка чтения адреса в режиме MEM.")
        End Function

        ''' <remarks>MEM mode write address 0: WR#=0, DS#/RD#=1, AS#/ADDR=0, D0-D7=output.</remarks>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341MemWriteAddr0(ByVal iIndex As Integer, ByVal iBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <remarks>MEM mode write address 1: WR#=0, DS#/RD#=1, AS#/ADDR=1, D0-D7=output.</remarks>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341MemWriteAddr1(ByVal iIndex As Integer, ByVal iBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Записывает данные в режиме MEM и возвращает число записанных байтов.
        ''' </summary>
        ''' <param name="dataToWrite">Данные для записи.</param>
        ''' <param name="addr0">Записываем в адрес A0 (true) или A1 (false). Выбор адреса позволяет запись в один из двух регистров.</param>
        Public Function MemWrite(ByVal dataToWrite As Byte(), Optional ByVal addr0 As Boolean = True) As Integer
            Dim res As Boolean = False
            Dim len As Integer = dataToWrite.Length
            If addr0 Then
                res = CH341MemWriteAddr0(DeviceIndex, dataToWrite, len)
            Else
                res = CH341MemWriteAddr1(DeviceIndex, dataToWrite, len)
            End If
            If res Then
                Return len
            End If
            Throw New Exception("Ошибка записи адреса в режиме MEM.")
        End Function

#End Region '/MEM

    End Class

End Namespace