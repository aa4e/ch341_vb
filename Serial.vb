Imports System.Runtime.InteropServices

Namespace Ch341

    ''' <summary>
    ''' Работа CH341 в режиме Serial.
    ''' </summary>
    Public MustInherit Class Serial
        Inherits Ch341Device

#Region "CTOR"

        Protected Sub New(index As Integer, Optional isExclusive As Boolean = False)
            MyBase.New(index, isExclusive)
        End Sub

#End Region '/CTOR

#Region "ENUMS"

        ''' <summary>
        ''' Режимы проверки данных в последовательном режиме.
        ''' </summary>
        Public Enum ParityModes As Integer
            None = 0
            Odd = 1
            Even = 2
            Mark = 3
            Space = 4
        End Enum

        ''' <summary>
        ''' Скорости передачи в последовательном режиме.
        ''' </summary>
        Public Enum BaudRates As Integer
            br50 = 50
            br75 = 75
            br100 = 100
            br110 = 110
            br134 = 134
            br150 = 150
            br300 = 300
            br600 = 600
            br900 = 900
            br1200 = 1200
            br1800 = 1800
            br2400 = 2400
            br3600 = 3600
            br4800 = 4800
            br9600 = 9600
            br14400 = 14400
            br19200 = 19200
            br28800 = 28800
            br33600 = 33600
            br38400 = 38400
            br56000 = 56000
            br57600 = 57600
            br76800 = 76800
            br115200 = 115200
            br128000 = 128000
            br153600 = 153600
            br230400 = 230400
            br460800 = 460800
            br921600 = 921600
            br1500000 = 1500000
            br2000000 = 2000000
        End Enum

#End Region '/ENUMS 

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341SetupSerial(iIndex As Integer, iParityMode As Integer, iBaudRate As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Задаёт параметры последовательного режима устройства. Работает только в последовательном режиме.
        ''' </summary>
        ''' <param name="parityMode">Задаёт режим проверки данных.</param>
        ''' <param name="baudRate">Скорость обмена, 50..3000000 бит/с.</param>
        Protected Sub SetupSerialMode(parityMode As ParityModes, baudRate As BaudRates)
            If (baudRate >= 50) AndAlso (baudRate <= 3000000) Then
                Dim res As Boolean = CH341SetupSerial(DeviceIndex, parityMode, baudRate)
                If res Then
                    Return
                End If
            End If
            Throw New Exception("Невозможно установить параметры последовательного режима устройства.")
        End Sub

        ''' <summary>
        ''' Задаёт параметры потокового режима последовательного порта.
        ''' </summary>
        ''' <param name="iMode">Задаёт режим:
        '''	биты 1..0 - скорость интерфейса I2C / частоту SCL: 00 = low speed / 20KHz, 01 = standard / 100KHz, 10 = fast / 400KHz, 11 = high speed / 750KHz.
        '''	бит 2 - соответствие и состояния портов ввода-вывода SPI: 0 = стандартный (D5 выход / D7 вход), 1 = с двумя линиями передачи (D5, D4 выход / D7, D6 вход);
        ''' бит 7 - порядок передачи чисел: 0 = LSB первый, 1 = MSB первый (стандартный).
        '''	Остальные биты заразервированы и должны быть 0.
        ''' </param>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Protected Friend Shared Function CH341SetStream(iIndex As Integer, iMode As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341SetDelaymS(iIndex As Integer, iDelay As Integer) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Задаёт аппаратную асинхронную задержку для потоковых операций.
        ''' CH341SetDelaymS sets the hardware asynchronous delay, returns quickly after the call, and delays the specified number of milliseconds before the next stream operation.
        ''' </summary>
        ''' <param name="delaMs">Задержка, мс.</param>
        Protected Sub SetDalay(delaMs As Integer)
            Dim res As Boolean = CH341SetDelaymS(DeviceIndex, delaMs)
            If res Then
                Return
            End If
            Throw New Exception("Ошибка выставления задержки.")
        End Sub

    End Class

End Namespace