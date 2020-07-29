Imports System.Runtime.InteropServices

Namespace Ch341

    ''' <summary>
    ''' Работа с CH341 в последовательном режиме SPI.
    ''' </summary>
    Public NotInheritable Class SpiMaster
        Inherits Serial

#Region "CTOR"

        ''' <summary>
        ''' Открывает устройство в режиме SPI.
        ''' </summary>
        ''' <param name="index">Индекс устройства в системе, начиная с 0.</param>
        ''' <param name="isExclusive">Включить ли эксклюзивное использование чипа.</param>
        ''' <param name="isDoubleInOut">Число входов-выходов SPI, 0 = single-input single-out (D5 out / D7 in), 1 = double-in and double-out (D4,D5 out / D6,D7 in)</param>
        ''' <param name="isMsbFirst">MSB первым (стандартный режим).</param>
        ''' <param name="wires">Число проводников режима SPI.</param>
        Public Sub New(index As Integer, Optional isExclusive As Boolean = False, Optional wires As WiresCountEnum = WiresCountEnum.Wires4, Optional isMsbFirst As Boolean = True, Optional isDoubleInOut As Boolean = False)
            MyBase.New(index, isExclusive)
            Me.WiresCount = wires
            SetInOutCount(isMsbFirst, isDoubleInOut)
        End Sub

#End Region '/CTOR

#Region "ENUMS"

        ''' <summary>
        ''' Число проводников режима SPI.
        ''' </summary>
        Public Property WiresCount As WiresCountEnum = WiresCountEnum.Wires4

        ''' <summary>
        ''' Число проводов, используемых в SPI.
        ''' </summary>
        Public Enum WiresCountEnum As Integer
            ''' <summary>
            ''' 3-проводной интерфейс SPI: тактовая частота на выводе DCK2/SCL, данные на DIO/SDA (квази-двунаправленный порт), выбор чипа D0/D1/D2. 
            ''' Скорость примерно 51 кслов/с.
            ''' </summary>
            Wires3
            ''' <summary>
            ''' 4-проводной интерфейс SPI: тактовая частота на выводе DCK/D3, выходные данные DOUT/D5, входные данные DIN/D7, выбор чипа D0/D1/D2. 
            ''' Скорость примерно 68 кбайт/с.
            ''' </summary>
            Wires4
            ''' <summary>
            ''' 5-проводной интерфейс SPI: Тактовая частота на выводе DCK/D3, выходные данные DOUT/D5 и DOUT2/D4, вхродные DIN/D7 b DIN2/D6, выбор чипа D0/D1/D2.
            ''' Скорость примерно 30 кбайт*2/с.
            ''' </summary>
            Wires5
        End Enum

#End Region '/ENUMS

        ''' <summary>
        ''' Задаёт число входов-выходов SPI, 0 = single-input single-out (D5 out / D7 in), 1 = double-in and double-out (D4,D5 out / D6,D7 in).
        ''' </summary>
        ''' <param name="isMsbFirst"></param>
        ''' <param name="isDoubleInOut"></param>
        Public Sub SetInOutCount(ByVal isMsbFirst As Boolean, ByVal isDoubleInOut As Boolean)
            Dim res As Boolean = False
            Dim mode As Integer = 0
            If isDoubleInOut Then
                mode = 3
                'WiresCount = WiresCountEnum.Wires5 'TEST
            End If
            If isMsbFirst Then
                mode = mode Or &H80
            End If
            res = CH341SetStream(DeviceIndex, mode)
            If res Then
                Return
            End If
            Throw New Exception("Ошибка выставления режима SPI.")
        End Sub

#Region "ЧТЕНИЕ И ЗАПИСЬ"

        ''' <summary>
        ''' Реализует обмен по 3-проводному интефрейсу SPI. 
        ''' Тактовая частота на выводе DCK2/SCL, данные на DIO/SDA (квази-двунаправленный ввод-вывод), выбор чипа D0/D1/D2, скорость примерно 51 кслов/с.
        ''' </summary>
        ''' <param name="chipSelect">Выбор ведомого: 
        ''' бит 7 = 0 - выбор чипа игнорируется, бит 7 = 1 - данные валидны; 
        ''' биты 1..0 = 00/01/10 соответственно активируют D0/D1/D2 низким уровнем.</param>
        ''' <param name="writeLength">Число байтов для передачи.</param>
        ''' <param name="ioBuffer">Буфер для записи в DIO, и сюда же сохраняются прочитанные с DIO данные.</param>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341StreamSPI3(index As Integer, chipSelect As Integer, writeLength As Integer, ioBuffer As Byte()) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Реализует обмен по 4-проводному интефрейсу SPI. 
        ''' Тактовая частота на выводе DCK/D3, выходные данные на DOUT/D5, входные данные на DIN/D7, выбор чипа D0/D1/D2, скорость примерно 68 кбайт/с.
        ''' </summary>
        ''' <param name="index">Индекс устройства.</param>
        ''' <param name="сhipSelect">Выбор ведомого: 
        ''' бит 7 = 0 - выбор чипа игнорируется, бит 7 = 1 - данные валидны; 
        ''' биты 1..0 = 00/01/10 соответственно активируют D0/D1/D2 низким уровнем.</param>
        ''' <param name="writeLength">Число байтов для передачи.</param>
        ''' <param name="ioBuffer">Буфер для записи в DOUT, и сюда же сохраняются прочитанные с DIN данные.</param>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341StreamSPI4(index As Integer, сhipSelect As Integer, writeLength As Integer, ioBuffer As Byte()) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary> 
        ''' Реализует обмен по 5-проводному интефрейсу SPI. 
        ''' Тактовая частота на выводе DCK/D3, выходные данные на DOUT/D5 и DOUT2/D4, входные на DIN/D7 и DIN2/D6, выбор чипа D0/D1/D2, скорость примерно 30*2 кбайт/с.
        ''' </summary>
        ''' <param name="chipSelect">Выбор ведомого: 
        ''' бит 7 = 0, вывод CS игнорируется, бит 7 = 1 данные валидны;
        ''' биты 1..0 = 00/01/10 соответственно активируют D0/D1/D2 низким уровнем.</param>
        ''' <param name="writeLength">Число байтов для передачи.</param>
        ''' <param name="ioBuffer">Буфер для записи в DOUT, и сюда же сохраняются прочитанные с DIN данные.</param>
        ''' <param name="ioBuffer2">Буфер для записи в DOUT2, и сюда же сохраняются прочитанные с DIN2 данные.</param>
        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341StreamSPI5(index As Integer, chipSelect As Integer, writeLength As Integer, ioBuffer As Byte(), ioBuffer2 As Byte()) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Записывает заданный массив и возвращает считанный массив по SPI.
        ''' </summary>
        ''' <param name="writeBuffer">Буфер для записи.</param>
        ''' <param name="readLength">Число байтов для чтения.</param>
        Public Function StreamSpi(writeBuffer As Byte(), readLength As Integer) As Byte()
            'TODO "Забил гвоздями" вывод CS. Надо реализовать возможность управлять им.
            Dim res As Boolean = False
            Dim writeLength As Integer = writeBuffer.Length
            ReDim Preserve writeBuffer(writeLength + readLength - 1)
            Select Case WiresCount
                Case WiresCountEnum.Wires3
                    res = CH341StreamSPI3(DeviceIndex, &H80, writeBuffer.Length, writeBuffer) 'TODO Ошибка передачи - разобраться.
                Case WiresCountEnum.Wires4
                    res = CH341StreamSPI4(DeviceIndex, &H80, writeBuffer.Length, writeBuffer)
                Case WiresCountEnum.Wires5
                    Dim readBuffer2(readLength - 1) As Byte
                    res = CH341StreamSPI5(DeviceIndex, &H80, writeBuffer.Length, writeBuffer, readBuffer2)
            End Select
            If res Then
                Return writeBuffer.Skip(writeLength).Take(readLength).ToArray()
            End If
            Throw New Exception("Ошибка чтения или записи массива байтов по SPI.")
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341BitStreamSPI(ByVal iIndex As Integer, ByVal iLength As Integer, ByVal ioBuffer As Byte()) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Управляет битовой передачей по SPI в 4- и 5-проводном режимах.
        ''' Тактовая частота на выводе DCK/D3, по умолчанию LOW, DOUT/D5 и DOUT2/D4 - выходы по срезу, DIN/D7 и DIN2/D6 по спаду. */
        ''' One byte in ioBuffer corresponds to D7-D0 pin, bit 5 is output to DOUT, bit 4 is output to DOUT2, bit 2..0 is output to D2-D0, bit 7 is input from DIN, bit 6 input from DIN2, bit 3 data ignored */
        ''' </summary>
        ''' <param name="writeBuffer"></param>
        ''' <param name="bitsToWrite">Число битов, которые следует передать, до 896 за раз. Не рекомендуется передавать более 256.</param>
        ''' <param name="readLength"></param>
        <Obsolete("Before calling this API, CH341Set_D5_D0 should be called to set the I/O direction of the D5-D0 pin of CH341 and set the default level of the pin")>
        Public Function BitSreamSpi(ByVal writeBuffer As Byte(), ByVal bitsToWrite As Integer, ByVal readLength As Integer) As Byte()
            'TODO Надо разобраться, перевод c китайского не точный.
            If (readLength > writeBuffer.Length) Then
                ReDim Preserve writeBuffer(readLength - 1)
            End If
            Dim res As Boolean = CH341BitStreamSPI(DeviceIndex, bitsToWrite, writeBuffer)
            If res Then
                Return writeBuffer
            End If
            Throw New Exception("Ошибка чтения или записи массива битов по SPI.")
        End Function

#End Region '/ЧТЕНИЕ И ЗАПИСЬ

    End Class

End Namespace