Imports System.Runtime.InteropServices
Imports System.Text

Namespace Ch341

    ''' <summary>
    ''' Работа с м/сх CH341.
    ''' </summary>
    Public Class Ch341Device
        Implements IDisposable

#Region "CONST"

        ''' <summary>
        ''' Путь к библиотеке CH341DLL.DLL.
        ''' </summary>
        Public Const DLL_PATH As String = "c:\Temp\CH341DLL.DLL"

        ''' <summary>
        ''' Неверный дескриптор.
        ''' </summary>
        Public Const INVALID_HANDLE_VALUE = -1

        ''' <summary>
        ''' Бесконечное время ожидания при обмене по USB.
        ''' </summary>
        Public Const USB_NO_TIMEOUT As UInteger = &HFFFFFFFFUI

#End Region '/CONST

#Region "ENUMS"

        ''' <summary>
        ''' Версии микросхемы CH341.
        ''' </summary>
        Public Enum IcVersion As Integer
            InvalidIc = 0
            Ch341 = &H10
            Ch341a = &H20
            Ch341a_ = &H30
        End Enum

        ''' <summary>
        ''' Состояние подключения/отключения устройства.
        ''' </summary>        
        ''' <remarks>
        ''' CH341_DEVICE_REMOVE = 0
        ''' CH341_DEVICE_REMOVE_PEND = 1
        ''' CH341_DEVICE_ARRIVAL = 3
        ''' </remarks>
        Public Enum DeviceEvent As Integer
            ''' <summary>
            ''' Device pullout event has been pulled out.
            ''' </summary>
            DeviceRemoved = 0
            ''' <summary>
            ''' The device is about to be pulled out.
            ''' </summary>
            DeviceRemovePending = 1
            ''' <summary>
            ''' Device insertion event, inserted.
            ''' </summary>
            DeviceInserted = 3
        End Enum

        ''' <summary>
        ''' Словарь типов м/сх.
        ''' </summary>
        Protected ReadOnly Property IcDict As New Dictionary(Of IcVersion, String) From {
            {IcVersion.InvalidIc, "Недействительный тип"},
            {IcVersion.Ch341, "CH341"},
            {IcVersion.Ch341a, "CH341A"},
            {IcVersion.Ch341a_, "CH341A"}
        }

#End Region '/ENUMS

#Region "STRUCTS"

        ''' <summary>
        ''' Define the WIN32 command interface structure.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure WIN32_COMMAND
            Public mFunction As Union1
            Public mLength As UInteger
            Public mBuffer As Union2

            <StructLayout(LayoutKind.Explicit)>
            Public Structure Union1
                <FieldOffset(0)> Public mFunction As UInteger
                <FieldOffset(0)> Public mStatus As Integer
            End Structure

            <StructLayout(LayoutKind.Explicit, CharSet:=CharSet.Ansi)>
            Public Structure Union2
                <FieldOffset(0)> Public mSetupPkt As USB_SETUP_PKT
                <FieldOffset(0)> <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)> Public mBuffer As Byte()
            End Structure
        End Structure

        ''' <summary>
        ''' USB Data request packet structure that controls the establishment phase of the transmission.
        ''' </summary>
        <StructLayout(LayoutKind.Sequential)>
        Public Structure USB_SETUP_PKT
            Public mUspReqType As Byte  ' 00H Request type
            Public mUspRequest As Byte  ' 01H Request code
            Public variable_1 As Union1 ' 02H-03H Value parameter
            Public variable_2 As Union2 ' 04H-05H Index parameter
            Public mLength As Short ' 06H-07H Data length of the data phase

            <StructLayout(LayoutKind.Explicit)>
            Public Structure Union1
                <FieldOffset(0)> Public variable_1 As Struct1
                <FieldOffset(0)> Public mUspValue As Short

                <StructLayout(LayoutKind.Sequential)>
                Public Structure Struct1
                    Public mUspValueLow As Byte  ' 02H Value parameter low byte
                    Public mUspValueHigh As Byte ' 03H Value parameter high byte
                End Structure
            End Structure

            <StructLayout(LayoutKind.Explicit)>
            Public Structure Union2
                <FieldOffset(0)> Public variable_1 As Struct1
                <FieldOffset(0)> Public mUspIndex As Short

                <StructLayout(LayoutKind.Sequential)>
                Public Structure Struct1
                    Public mUspIndexLow As Byte  ' 04H Index parameter low byte
                    Public mUspIndexHigh As Byte ' 05H Index parameter high byte
                End Structure
            End Structure
        End Structure

#End Region '/STRUCTS

#Region "READ-ONLY PROPS"

        ''' <summary>
        ''' Индекс устройства в системе.
        ''' </summary>
        Public ReadOnly Property DeviceIndex As Integer
            Get
                Return _DeviceIndex
            End Get
        End Property
        Private _DeviceIndex As Integer = INVALID_HANDLE_VALUE

        ''' <summary>
        ''' Дескриптор устройства.
        ''' </summary>
        Public ReadOnly Property Handle As IntPtr
            Get
                Return _Handle
            End Get
        End Property
        Private _Handle As IntPtr = IntPtr.Zero

        ''' <summary>
        ''' Открыто ли устройство.
        ''' </summary>
        Public ReadOnly Property IsOpened As Boolean
            Get
                Return (Handle <> IntPtr.Zero)
            End Get
        End Property

        ''' <summary>
        ''' Состояние вывода ERR# (pin 5, Input) - IN0.
        ''' </summary>
        Public ReadOnly Property PinErrState As Boolean
            Get
                Return ((GetInput() >> 8 And 1) = 1)
            End Get
        End Property

        ''' <summary>
        ''' Состояние вывода PEMP (pin 6, Input) - IN1.
        ''' </summary>
        Public ReadOnly Property PinPempState As Boolean
            Get
                Return ((GetInput() >> 9 And 1) = 1)
            End Get
        End Property

        ''' <summary>
        ''' Состояние вывода INT# (pin 7, Input) - Interrupt request (по переднему фронту).
        ''' </summary>
        Public ReadOnly Property PinIntState As Boolean
            Get
                Return ((GetInput() >> 10 And 1) = 1)
            End Get
        End Property

        ''' <summary>
        ''' Состояние вывода SEL (pin 8, Input) - IN3.
        ''' </summary>
        Public ReadOnly Property PinSelState As Boolean
            Get
                Return ((GetInput() >> 11 And 1) = 1)
            End Get
        End Property

        ''' <summary>
        ''' Состояние вывода SDA (pin 23).
        ''' </summary>
        Public ReadOnly Property PinSdaState As Boolean
            Get
                Return ((GetInput() >> 23 And 1) = 1)
            End Get
        End Property

        ''' <summary>
        ''' Состояние вывода BUSY/WAIT# (pin 27, Input) - /Wait.
        ''' </summary>
        Public ReadOnly Property PinWaitState As Boolean
            Get
                Return ((GetInput() >> 13 And 1) = 1)
            End Get
        End Property

        ''' <summary>
        ''' Состояние вывода AUTO341#/DATAS# (pin 4, Output) - /Data Select.
        ''' </summary>
        Public ReadOnly Property PinDataStrobeState As Boolean
            Get
                Return ((GetInput() >> 14 And 1) = 1)
            End Get
        End Property

        ''' <summary>
        ''' Состояние вывода SLCTIN#/ADDRS# (pin 3, Output) - /Address select.
        ''' </summary>
        Public ReadOnly Property PinAddrStrobeState As Boolean
            Get
                Return ((GetInput() >> 15 And 1) = 1)
            End Get
        End Property

        ''' <summary>
        ''' Состояние линий D0...D7 (двунаправленные выводы 15..22) - 8-битная шина данных.
        ''' </summary>
        Public ReadOnly Property PinDataState As Boolean()
            Get
                Dim d As New BitArray({GetInput()})
                Return {d(0), d(1), d(2), d(3), d(4), d(5), d(6), d(7)}
            End Get
        End Property

#End Region '/READ-ONLY PROPS

#Region "ИНФОРМАЦИОННЫЕ МЕТОДЫ"

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341GetVersion() As Integer
        End Function

        ''' <summary>
        ''' Возвращает версию библиотеки.
        ''' </summary>
        Public Shared Function GetDllVersion() As Integer
            Dim version As Integer = CH341GetVersion()
            Return version
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341GetDrvVersion() As Integer
        End Function

        ''' <summary>
        ''' Возвращает версию драйвера.
        ''' </summary>
        Public Function GetDriverVersion() As Integer
            Dim drvVer As Integer = CH341GetDrvVersion()
            Return drvVer
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341GetVerIC(iIndex As Integer) As Integer
        End Function

        ''' <summary>
        ''' Возвращает тип микросхемы CH341.
        ''' </summary>
        Public Function GetIcVersion() As IcVersion
            Dim verIc As IcVersion = CType(CH341GetVerIC(DeviceIndex), IcVersion)
            Return verIc
        End Function

        ''' <summary>
        ''' Возвращает описание типа микросхемы CH341.
        ''' </summary>
        Public Function GetIcVersionReadable() As String
            Dim ver As IcVersion = GetIcVersion()
            Dim s As String = IcDict(ver)
            Return s
        End Function

        ''' <param name="oBuffer">Points to a buffer large enough to hold the descriptor.</param>
        ''' <param name="ioLength">Points to the length unit, the length to be read when input, and the length to be read after returning.</param>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341GetDeviceDescr(iIndex As Integer, oBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Возвращает описатель устройства.
        ''' </summary>
        Public Function GetDeviceDescriptor() As String
            Dim len As Integer = &H400
            Dim descr(len - 1) As Byte
            Dim res As Boolean = CH341GetDeviceDescr(DeviceIndex, descr, len)
            Dim s As String = String.Empty
            If res AndAlso (len > 0) Then
                'Dim chinaEnc As Encoding = Encoding.GetEncoding(65000) '"Unicode")
                's = chinaEnc.GetString(descrPtr, 0, len)
                s = Encoding.Unicode.GetString(descr, 0, len)
            End If
#If DEBUG Then
            For i As Integer = 0 To len - 1
                Console.Write(descr(i).ToString("X2")) '.PadLeft(2, "0"c))
                Console.Write(" ")
            Next
            Console.WriteLine()
#End If
            Return s
        End Function

        ''' <remarks>Return a buffer pointing to the CH341 device name, returning NULL if an error occurs.</remarks>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341GetDeviceName(iIndex As Integer) As IntPtr
        End Function

        ''' <summary>
        ''' Возвращает название устройства (путь в системе).
        ''' </summary>
        Public Function GetDeviceName() As String
            Dim ptr As IntPtr = CH341GetDeviceName(DeviceIndex)
            Dim s As String = Marshal.PtrToStringAnsi(ptr)
            Return s
        End Function

        ''' <param name="oBuffer">Points to a buffer large enough to hold the descriptor.</param>
        ''' <param name="ioLength">Points to the length unit, the length to be read when input, and the length to be read after returning.</param>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341GetConfigDescr(iIndex As Integer, oBuffer As Byte(), ByRef ioLength As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Возвращает описатель конфигурации.
        ''' </summary>
        Public Function GetConfigDescriptor() As String
            Dim len As Integer = &H400
            Dim descr(len - 1) As Byte
            Dim res As Boolean = CH341GetConfigDescr(DeviceIndex, descr, len)
            Dim s As String = String.Empty
            If res AndAlso (len > 0) Then
                s = Encoding.Unicode.GetString(descr, 0, len)
            End If
#If DEBUG Then
            For i As Integer = 0 To len - 1
                Console.Write(descr(i).ToString("X2")) '.PadLeft(2, "0"c))
                Console.Write(" ")
            Next
            Console.WriteLine()
#End If
            Return s
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341DriverCommand(iIndex As Integer, ByRef ioCommand As WIN32_COMMAND) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Directly pass the command to the driver, return 0 if error occurs, otherwise return the data length.
        ''' </summary>
        ''' <remarks>
        ''' The program returns the data length after the call, and still returns the command structure. 
        ''' If it is a read operation, the data is returned in the command structure.
        ''' The returned data length is 0 when the operation fails. 
        ''' When the operation succeeds, it is the length of the entire command structure. 
        ''' For example, if one byte is read, mWIN32_COMMAND_HEAD+1 is returned.
        ''' The command structure is provided before the call: pipe number or command function code, length of access data (optional), data (optional)
        ''' After the command structure is called, it returns: the operation status code, the length of the subsequent data (optional),
        ''' operation status code is the code defined by WINDOWS, you can refer to NTSTATUS.H,
        ''' The length of the subsequent data refers to the length Of the data returned by the read operation, the data is stored in the subsequent buffer, and is generally 0 for write operations.
        ''' </remarks>
        <Obsolete("Пока не работает - Неверно упакован тип WIN32_COMMAND.")>
        Public Function DriverCommand() As IntPtr 'WIN32_COMMAND
            ''Dim ptr As IntPtr = IntPtr.Zero
            'Dim ic As New WIN32_COMMAND 'IntPtr.Zero ' 
            'Dim res As Boolean = CH341DriverCommand(DeviceIndex, ic)
            ''Return ic
            Return IntPtr.Zero
        End Function

#End Region '/ИНФОРМАЦИОННЫЕ МЕТОДЫ

#Region "ОТКРЫТИЕ, ЗАКРЫТИЕ, СБРОС УСТРОЙСТВА"

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341OpenDevice(index As Integer) As IntPtr
        End Function

        ''' <summary>
        ''' Открывает устройство с заданным индексом.
        ''' </summary>
        ''' <param name="index">Индекс устройства в системе, начиная с 0.</param>
        ''' <param name="exclusive">Открыть эксклюзивно или разрешить другим процессам использовать устройство.</param>
        Protected Sub New(index As Integer, Optional exclusive As Boolean = False)
            _DeviceIndex = index
            _Handle = CH341OpenDevice(index)
            If (Handle.ToInt32() = INVALID_HANDLE_VALUE) OrElse (Handle = IntPtr.Zero) Then
                Throw New SystemException(String.Format("Невозможно открыть устройство {0}.", index))
            End If
            SetExclusive(exclusive)
            Dim notifDeleg As New mPCH341_NOTIFY_ROUTINE(AddressOf UsbEventHandler)
            SetDeviceNotify(notifDeleg)
        End Sub

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Sub CH341CloseDevice(index As Integer)
        End Sub

        ''' <summary>
        ''' Закрывает устройство.
        ''' </summary>
        Public Sub CloseDevice()
            If IsOpened Then
                CancelDeviceNotify()
                CH341CloseDevice(DeviceIndex)
                _DeviceIndex = INVALID_HANDLE_VALUE
                _Handle = IntPtr.Zero
            End If
        End Sub

        ''' <summary>
        ''' Закрывает указанное устройство.
        ''' </summary>
        ''' <param name="index">Индекс устройства в системе, начиная с 0.</param>
        Public Shared Sub CloseDevice(index As Integer)
            CH341CloseDevice(index)
        End Sub

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341ResetDevice(index As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Сбрасывает устройство (возвращает к исходному состоянию после загрузки).
        ''' </summary>
        Public Sub ResetDevice()
            If CH341ResetDevice(DeviceIndex) Then
                Return
            End If
            Throw New Exception("Ошибка сброса устройства.")
        End Sub

        ''' <summary>
        ''' Сбрасывает заданное устройство.
        ''' </summary>
        ''' <param name="index">Индекс устройства в системе, начиная с 0.</param>
        ''' <remarks>Статический метод добавлен, чтобы сбрасывать устройство в ситуации "зависания".</remarks>
        Public Shared Sub ResetDevice(index As Integer)
            If CH341ResetDevice(index) Then
                Return
            End If
            Throw New Exception("Ошибка сброса устройства.")
        End Sub

#End Region '/ОТКРЫТИЕ, ЗАКРЫТИЕ, СБРОС УСТРОЙСТВА

#Region "РАБОТА С ПРЕРЫВАНИЯМИ"

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341SetIntRoutine(index As Integer, cb As mPCH341_INT_ROUTINE) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Обработчик прерывания.
        ''' </summary>
        ''' <param name="status">
        ''' Данные о статусе:
        ''' биты 0..7 соответствуют выводам D0..D7,
        ''' бит 8 - ERR#,
        ''' бит 9 - PEMP, 
        ''' бит 10 - INT#, 
        ''' бит 11 - SLCT.
        ''' </param>
        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Public Delegate Sub mPCH341_INT_ROUTINE(status As Integer)

        ''' <summary>
        ''' Задаёт обработчик прерывания.
        ''' </summary>
        ''' <param name="routineHandler">Обработчик прерывания.</param>
        Public Sub SetInterruptRoutine(routineHandler As mPCH341_INT_ROUTINE)
            If CH341SetIntRoutine(DeviceIndex, routineHandler) Then
                Return
            End If
            Throw New Exception("Ошибка установки обработчика прерывания.")
        End Sub

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341ReadInter(index As Integer, ByRef status As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Читает данные прерывания.
        ''' </summary>
        ''' <remarks>
        ''' биты 0..7 соответствуют состоянию пинов D0..D7,
        ''' бит 8  - ERR#,
        ''' бит 9  - PEMP, 
        ''' бит 10 - INT#,
        ''' бит 11 - SLCT.
        ''' </remarks>
        <Obsolete("Вызывает зависание, разобраться.")>
        Public Function ReadInterrupt() As Integer
            Dim s As Integer = 0
            If CH341ReadInter(DeviceIndex, s) Then
                Return s
            End If
            Throw New Exception("Ошибка чтения прерывания.")
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341ResetInter(index As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Сбрасывает прерывание операции чтения.
        ''' </summary>
        Public Sub ResetInterrupt()
            If CH341ResetInter(DeviceIndex) Then
                Return
            End If
            Throw New Exception("Ошибка сброса прерывания.")
        End Sub

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341AbortInter(index As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Отмена операции чтения прерывания.
        ''' </summary>
        <Obsolete("Вызывает зависание.")>
        Public Sub AbortInterrupt()
            If CH341AbortInter(DeviceIndex) Then
                Return
            End If
            Throw New Exception("Ошибка отмены прерывания.")
        End Sub

#End Region '/РАБОТА С ПРЕРЫВАНИЯМИ

#Region "ПАРАМЕТРЫ USB"

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341SetTimeout(index As Integer, writeTimeout As UInteger, readTimeout As UInteger) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Задаёт время ожидания для операций чтения и записи по USB.
        ''' </summary>
        ''' <param name="writeTimeoutUs">Задаёт время ожидания при записи в мс. По умолчанию <see cref="USB_NO_TIMEOUT "/> - без таймаута.</param>
        ''' <param name="readTimeoutUs">Задаёт время ожидания при чтении в мс. По умолчанию <see cref="USB_NO_TIMEOUT "/> - без таймаута.</param>
        Public Sub SetUsbTimeout(Optional writeTimeoutUs As UInteger = USB_NO_TIMEOUT, Optional readTimeoutUs As UInteger = USB_NO_TIMEOUT)
            If (writeTimeoutUs > 0) AndAlso (writeTimeoutUs <= USB_NO_TIMEOUT) AndAlso (readTimeoutUs > 0) AndAlso (readTimeoutUs <= USB_NO_TIMEOUT) Then
                If CH341SetTimeout(DeviceIndex, writeTimeoutUs, readTimeoutUs) Then
                    Return
                End If
                Throw New Exception("Ошибка установки времени ожидания USB.")
            End If
        End Sub

        ''' <summary>
        ''' Событие при изменении состояния устройства.
        ''' </summary>
        ''' <param name="evt">Состояние устройства.</param>
        Public Event DeviceStateEvent(evt As DeviceEvent)

        Private Sub UsbEventHandler(evt As DeviceEvent)
            RaiseEvent DeviceStateEvent(evt)
        End Sub

        ''' <summary>
        ''' Делегат оповещения о событии.
        ''' </summary>
        <UnmanagedFunctionPointer(CallingConvention.StdCall)>
        Private Delegate Sub mPCH341_NOTIFY_ROUTINE(eventStatus As DeviceEvent) ', msg As Integer, wParam As Integer, lParam As Integer)

        ''' <param name="deviceID">Необязательный параметр. Указатель на строку, задающую ID устройства, которое нужно мониторить. Строка заканчивается \0.</param>
        ''' <param name="notifyRoutine">Адрес функции. Задаёт программу обратного вызова, которая выполняется, когда событие зарегистрировано. NULL отменяет уведомления о событиях.</param>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341SetDeviceNotify(index As Integer, <MarshalAs(UnmanagedType.LPStr)> deviceID As StringBuilder, notifyRoutine As mPCH341_NOTIFY_ROUTINE) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Задаёт обработчик уведомлений о событиях подключения и отключения устройства.
        ''' </summary>
        ''' <param name="callback">Адрес функции. Задаёт программу обратного вызова, которая выполняется, когда событие зарегистрировано. </param>
        Private Sub SetDeviceNotify(callback As mPCH341_NOTIFY_ROUTINE)
            If CH341SetDeviceNotify(DeviceIndex, Nothing, callback) Then
                Return
            End If
            Throw New Exception("SetDeviceNotify()")
        End Sub

        ''' <summary>
        ''' Отменяет уведомление о событиях подключения и отключения устройства.
        ''' </summary>
        Private Sub CancelDeviceNotify() 'TEST
            'Dim devId As New StringBuilder(vbNullString)
            If CH341SetDeviceNotify(DeviceIndex, Nothing, Nothing) Then
                Return
            End If
            Throw New Exception("CancelDeviceNotify()")
        End Sub

#End Region '/ПАРАМЕТРЫ USB

#Region "УПРАВЛЕНИЕ"

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341SetExclusive(iIndex As Integer, <MarshalAs(UnmanagedType.Bool)> iExclusive As Boolean) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Задаёт или снимает эксклюзивное использование устройства (разрешает/запрещает другим процессам подключаться к устройству, пока оно открыто).
        ''' </summary>
        ''' <param name="isExclusive">True - устройство используется эксклюзивно (не даёт переоткрывать себя), False - может быть открыто несколько экземпляров устройства.</param>
        Public Sub SetExclusive(isExclusive As Boolean)
            If CH341SetExclusive(DeviceIndex, isExclusive) Then
                Return
            End If
            Throw New Exception("Ошибка выставления эксклюзивного режима работы с устройством.")
        End Sub

#End Region '/УПРАВЛЕНИЕ

#Region "РАБОТА С ВНУТРЕННИМ БУФЕРОМ"

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341FlushBuffer(iIndex As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Очищает буфер.
        ''' </summary>
        Public Sub ClearBuffer()
            If CH341FlushBuffer(DeviceIndex) Then
                Return
            End If
            Throw New Exception("Ошибка очистки буфера устройства.")
        End Sub

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341QueryBufUpload(iIndex As Integer) As Integer
        End Function

        ''' <summary>
        ''' Возвращает количество находящихся во внутреннем исходящем буфере устройства пакетов данных. Если буфер не подключён, возвращается -1.
        ''' </summary>
        Public Function GetUploadBuffer() As Integer
            Dim packetsCount As Integer = CH341QueryBufUpload(DeviceIndex)
            Return packetsCount
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341QueryBufDownload(iIndex As Integer) As Integer
        End Function

        ''' <summary>
        ''' Возвращает количество находящихся во внутреннем входящем буфере устройства данных (ещё не отправленных). Если буфер не подключён, возвращается -1.
        ''' </summary>
        Public Function GetDownloadBuffer() As Integer
            Dim packetsCount As Integer = CH341QueryBufDownload(DeviceIndex)
            Return packetsCount
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341SetBufUpload(iIndex As Integer, <MarshalAs(UnmanagedType.Bool)> iEnableOrClear As Boolean) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Включает или выключает режим внутреннего исходящего буфера при обмене с устройством по USB.
        ''' </summary>
        ''' <param name="enableBuffer">True - включает внутренний буфер и очищает все данные в буфере, False - отключает режим буфера и переводит устройство в режим прямой выгрузки.
        ''' Чтение с помощью вызова <see cref="Parallel.ReadData"/> немедленно возвращает имеющиеся в буфере данные.
        ''' </param>
        Public Sub SetUploadBuffer(enableBuffer As Boolean)
            If CH341SetBufUpload(DeviceIndex, enableBuffer) Then
                Return
            End If
            Throw New Exception("Ошибка выставления исходящего буфера.")
        End Sub

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341SetBufDownload(iIndex As Integer, <MarshalAs(UnmanagedType.Bool)> iEnableOrClear As Boolean) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Включает или выключает режим внутреннего входящего буфера при обмене с устройством по USB.
        ''' </summary>
        ''' <param name="enableBuffer">True - включает внутренний буфер и очищает все данные в буфере, False - отключает режим буфера и переводит устройство в режим прямой загрузки.
        ''' Запись с помощью вызова <see cref="Parallel.WriteData"/> немедленно отправляет все имеющиеся в буфере данные.
        ''' </param>
        Public Sub SetDownloadBuffer(enableBuffer As Boolean)
            If CH341SetBufDownload(DeviceIndex, enableBuffer) Then
                Return
            End If
            Throw New Exception("Ошибка выставления входящего буфера.")
        End Sub

#End Region '/РАБОТА С ВНУТРЕННИМ БУФЕРОМ

#Region "ВЫБОРОЧНАЯ КОНФИГУРАЦИЯ ВЫВОДОВ"

        'Pin assignment CH341A:
        'Bi-directional pins (input or output)
        '   Bit 0-7 = D0-D7 pin 15-22
        '   Bit 8 = ERR# pin 5
        '   Bit 9 = PEMP pin 6
        '   Bit 10 = ACK pin 7
        '   Bit 11 = SLCT pin 8
        '   Bit 12 = unused -
        '   Bit 13 = BUSY/WAIT# pin 27
        '   Bit 14 = AUTOFD#/DATAS# pin 4
        '   Bit 15 = SLCTIN#/ADDRS# pin 3
        'Pseudo-bi-directional pins (Combination of input And open-collector output) :
        '   Bit 23 = SDA pin 23 (for functions GetInput/GetStatus)
        '   Bit 16 = SDA pin 23 (for funktion SetOutput)
        'Uni-directional pins (output only):
        '   Bit 16 = RESET# pin 26
        '   Bit 17 = WRITE# pin 25
        '   Bit 18 = SCL pin pin 24

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341GetStatus(iIndex As Integer, ByRef iStatus As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Возвращает состояния портов ввода-вывода данных. Лучше использовать более эффективный метод <see cref="GetInput"/>.
        ''' </summary>
        ''' <remarks>
        ''' Биты 0...7 отвечают за состояния пинов D0..D7,
        ''' бит 8 - ERR# - пин 5,
        ''' бит 9 - PEMP - пин 6,
        ''' бит 10 - INT# - пин 7,
        ''' бит 11 - SLCT - пин 8,
        ''' бит 13 - BUDY/WAIT# - пин 27,
        ''' бит 14 - AUTO341#/DATAS# - пин 4,
        ''' бит 15 - SLCTIN#/ADDRS# - пин 3,
        ''' бит _23_ - SDA - пин 23.
        ''' </remarks>
        Public Function GetPinsState() As Integer
            Dim stat As Integer = 0
            If CH341GetStatus(DeviceIndex, stat) Then
                Return stat
            End If
            Throw New Exception("Невозможно получить состояние портов ввода-вывода.")
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341GetInput(iIndex As Integer, ByRef iStatus As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Возвращает состояния портов ввода-вывода данных и статуса устройства. Эффективность выше, чем у метода <see cref="GetPinsState()"/>.
        ''' </summary>
        ''' <remarks>
        ''' Биты 0...7 отвечают за состояния пинов D0..D7,
        ''' бит 8 - ERR# - пин 5,
        ''' бит 9 - PEMP - пин 6,
        ''' бит 10 - INT# - пин 7,
        ''' бит 11 - SLCT - пин 8,
        ''' бит 13 - BUDY/WAIT# - пин 27,
        ''' бит 14 - AUTO341#/DATAS# - пин 4,
        ''' бит 15 - SLCTIN#/ADDRS# - пин 3,
        ''' бит _23_ - SDA - пин 23.
        ''' </remarks>
        Public Function GetInput() As Integer
            Dim status As Integer = 0
            If CH341GetInput(DeviceIndex, status) Then
                Return status
            End If
            Throw New Exception("Невозможно получить состояние портов ввода-вывода.")
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341SetOutput(ByVal iIndex As Integer, ByVal iEnable As Integer, ByVal iSetDirOut As Integer, ByVal iSetDataOut As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Задаёт направление и состояния пинов ввода-вывода.
        ''' </summary>
        ''' <param name="dataValid">Состояния валидности данных:
        ''' - бит 0 - 1 показывает, что биты 15..8  <paramref name="pinsState"/> валидны, в противном случае игнорируются;
        '''	- бит 1 - 1 показывает, что биты 15..8 <paramref name="pinsDirection"/> валидны, в противном случае игнорируются;
        '''	- бит 2 - 1 показывает, что биты 7..0 <paramref name="pinsState"/> валидны, в противном случае игнорируются;
        '''	- бит 3 - 1 показывает, что биты 7..0 <paramref name="pinsDirection"/> валидны, в противном случае игнорируются;
        '''	- бит 4 - 1 показывает, что биты 23..16 <paramref name="pinsState"/> валидны, в противном случае игнорируются.
        ''' </param>
        ''' <param name="pinsDirection">Задаёт направления портов I/O. Бит 0 соответствует входу, бит 1 - выходу (осторожно!). Значение по умолчанию для параллельного порта 0x000FC000.</param>
        ''' <param name="pinsState">Задаёт состояния портов ввода-вывода. Бит 0 соответствует низкому уровню, 1 - высокому: 
        '''	- биты 7..0 соответствуют пинам D7..D0;
        '''	- бит 8  - ERR#; 
        '''	- бит 9  - PEMP; 
        '''	- бит 10 - INT#; 
        '''	- бит 11 - SLCT; 
        '''	- бит 13 - WAIT#; 
        '''	- бит 14 - DATAS#/READ#; 
        '''	- бит 15 - ADDRS#/ADDR/ALE.
        '''	Следующие пины могут быть только выходами, независимо от заданного направления: 
        '''	- бит 16 соответствует пину RESET#; 
        '''	- бит 17 - WRITE#; 
        '''	- бит 18 - SCL; 
        '''	- бит 29 - SDA.
        ''' </param>
        ''' <remarks>
        ''' ***** Использовать этот API с осторожностью, т.к. неверное задание состояний и направлений выводов может привести к короткому замыканию и выходу микросхемы из строя! *****
        ''' Пример: 
        ''' CH341SetOutput(0, $FF, $FF, $F0) задаёт пины D0..D7 как выходы. Остальные линии не затрагиваются. D0...D3 - в состоянии LOW, D4...D7 - в состоянии HIGH.
        ''' </remarks>
        <Obsolete("Использовать с осторожностью, т.к. неверное задание состояний и направлений выводов может привести к короткому замыканию и выходу микросхемы из строя!")>
        Public Sub SetIoPinsDirection(dataValid As Integer, pinsDirection As Integer, pinsState As Integer)
            If CH341SetOutput(DeviceIndex, dataValid, pinsDirection, pinsState) Then
                Return
            End If
            Throw New Exception("Не удалось задать направления и состояния выводов.")
        End Sub

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341Set_D5_D0(iIndex As Integer, iSetDirOut As Integer, iSetDataOut As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Задаёт направления и состояния портов ввода-выода D0..D5. Эффективнее, чем <see cref="SetIOPinsDirection"/>.
        ''' </summary>
        ''' <param name = "pinsDirection">Задаёт направления портов D0..D5. Бит 0 соответствует входу, бит 1 - выходу (осторожно!). Значение по умолчанию для параллельного порта 0x00.</param>
        ''' <param name="pinsState">Биты 0..5 задают состояния портов D0..D5. Если выводы в режиме выхода, бит 0 соответствует низкому уровню, 1 - высокому.</param>
        <Obsolete("Использовать с осторожностью, т.к. неверное задание состояний и направлений выводов может привести к короткому замыканию и выходу микросхемы из строя!")>
        Public Sub SetDataPinsDirection(pinsDirection As Integer, pinsState As Integer)
            If CH341Set_D5_D0(DeviceIndex, pinsDirection, pinsState) Then
                Return
            End If
            Throw New Exception("Не удалось задать направления и состояния выводов D0..D5")
        End Sub

#End Region '/ВЫБОРОЧНАЯ КОНФИГУРАЦИЯ ВЫВОДОВ

#Region "IDisposable Support"

        Private DisposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If (Not DisposedValue) Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    CloseDevice()
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            DisposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            ' TODO: uncomment the following line if Finalize() is overridden above.
            ' GC.SuppressFinalize(Me)
        End Sub

#End Region

    End Class '/Ch341Device

End Namespace