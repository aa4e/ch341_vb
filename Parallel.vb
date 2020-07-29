Imports System.Runtime.InteropServices

Namespace Ch341

    ''' <summary>
    ''' Общие для параллельных режимов микросхемы CH341.
    ''' </summary>
    Public MustInherit Class Parallel
        Inherits Ch341Device

#Region "CTOR"

        Protected Sub New(index As Integer, Optional isExclusive As Boolean = False)
            MyBase.New(index, isExclusive)
        End Sub

#End Region '/CTOR

#Region "ENUM"

        ''' <summary>
        ''' Параллельные режимы работы.
        ''' </summary>
        Public Enum ParallelMode As Integer
            EPP17 = 0
            EPP19 = 1
            MEM = 2
            KeepCurrent = 8 'любое число > 4
        End Enum

#End Region '/ENUM

#Region "УПРАВЕЛНИЕ"

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341InitParallel(iIndex As Integer, iMode As ParallelMode) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Сбрасывает устройство и инициализирует параллельный порт. На ножку RST# подаётся нулевой импульс.
        ''' </summary>
        ''' <param name="mode">Задаёт режим работы параллельного порта: 0 - EPP / EPP версии 1.7, 1 - EPP версии 1.9, 2 - режим MEM, >= 0x00000100 - сохранять текущий режим.</param>
        Protected Friend Sub InitParallel(Optional mode As ParallelMode = ParallelMode.EPP17)
            Dim res As Boolean = CH341InitParallel(DeviceIndex, mode)
            If res Then
                Return
            End If
            Throw New Exception(String.Format("Ошибка инициализации параллельного режима {0}.", mode))
        End Sub

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341SetParaMode(iIndex As Integer, iMode As ParallelMode) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Устанавливает настройки режима параллельного порта.
        ''' </summary>
        ''' <param name="mode">Задаёт режим работы параллельного порта: 0 - EPP / EPP версии 1.7, 1 - EPP версии 1.9, 2 - режим MEM.</param>
        Protected Sub SetParallelMode(Optional mode As ParallelMode = ParallelMode.EPP17)
            Dim res As Boolean = CH341SetParaMode(DeviceIndex, mode)
            If res Then
                Return
            End If
            Throw New Exception("Ошибка выставления параллельного режима.")
        End Sub

#End Region '/УПРАВЕЛНИЕ

#Region "ЧТЕНИЕ ДАННЫХ EPP/MEM"

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341ReadData0(iIndex As Integer, oBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341ReadData1(iIndex As Integer, oBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Читает блок данных из порта 0 или 1.
        ''' </summary>
        ''' <param name="lenToRead">Сколько байтов прочитать из порта.</param>
        ''' <param name="port0">Читаем из порта 0 (true) или 1 (false).</param>
        Protected Function ReadData(lenToRead As Integer, Optional port0 As Boolean = True) As Byte()
            Dim readBuffer(lenToRead - 1) As Byte
            Dim len As Integer = lenToRead
            Dim res As Boolean = False
            If port0 Then
                res = CH341ReadData0(DeviceIndex, readBuffer, len)
            Else
                res = CH341ReadData1(DeviceIndex, readBuffer, len)
            End If
            If res Then
                Return readBuffer
            End If
            Throw New Exception("Ошибка чтения из параллельного порта.")
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341AbortRead(iIndex As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Отменяет операцию чтения.
        ''' </summary>
        Protected Sub AbortRead()
            Dim res As Boolean = CH341AbortRead(DeviceIndex)
            If res Then
                Return
            End If
            Throw New Exception("Ошибка отмены чтения.")
        End Sub

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341ResetRead(iIndex As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Сбрасывает операцию чтения блока данных.
        ''' </summary>
        Protected Sub ResetRead()
            Dim res As Boolean = CH341ResetRead(DeviceIndex)
            If res Then
                Return
            End If
            Throw New Exception("Ошибка сброса операции чтения.")
        End Sub

#End Region '/ЧТЕНИЕ ДАННЫХ EPP/MEM

#Region "ЗАПИСЬ ДАННЫХ EPP/MEM"

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341WriteData0(iIndex As Integer, iBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341WriteData1(iIndex As Integer, iBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Записывает массив данных в заданный порт и возвращает число записанных байтов.
        ''' </summary>
        ''' <param name="dataToWrite">Массив данных.</param>
        ''' <param name="port0">Записываем в порт 0 (true) или 1 (false).</param>
        Protected Function WriteData(dataToWrite As Byte(), Optional port0 As Boolean = True) As Integer
            Dim wroteLen As Integer = dataToWrite.Length
            Dim res As Boolean = False
            If port0 Then
                res = CH341WriteData0(DeviceIndex, dataToWrite, wroteLen)
            Else
                res = CH341WriteData1(DeviceIndex, dataToWrite, wroteLen)
            End If
            If res Then
                Return wroteLen
            End If
            Throw New Exception("Ошибка записи в параллельный порт.")
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341AbortWrite(iIndex As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Отменяет операцию записи.
        ''' </summary>
        Protected Sub AbortWrite()
            Dim res As Boolean = CH341AbortWrite(DeviceIndex)
            If res Then
                Return
            End If
        End Sub

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341ResetWrite(iIndex As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Сбрасывает операцию записи.
        ''' </summary>
        Protected Sub ResetWrite()
            Dim res As Boolean = CH341ResetWrite(DeviceIndex)
            If res Then
                Return
            End If
            Throw New Exception("Ошибка сброса операции записи.")
        End Sub

#End Region '/ЗАПИСЬ ДАННЫХ EPP/MEM

    End Class

End Namespace