Imports System.Runtime.InteropServices

Namespace Ch341

    ''' <summary>
    ''' Работа CH341 в последовательном режиме IIC. CH341 всегда мастер.
    ''' </summary>
    Public NotInheritable Class I2cMaster
        Inherits Serial

#Region "CTOR"

        ''' <summary>
        ''' Открывает I2C устройство с заданной скоростью шины.
        ''' </summary>
        ''' <param name="index"></param>
        ''' <param name="isExclusive"></param>
        ''' <param name="speed"></param>
        Public Sub New(index As Integer, Optional isExclusive As Boolean = False, Optional speed As I2cSpeed = I2cSpeed.Standard)
            MyBase.New(index, isExclusive)
            SetSpeed(speed)
        End Sub

#End Region '/CTOR

#Region "ENUMS"

        ''' <summary>
        ''' Типы ПЗУ.
        ''' </summary>
        Public Enum EepromTypes As Integer
            ID_24C01 = 0
            ID_24C02 = 1
            ID_24C04 = 2
            ID_24C08 = 3
            ID_24C16 = 4
            ID_24C32 = 5
            ID_24C64 = 6
            ID_24C128 = 7
            ID_24C256 = 8
            ID_24C512 = 9
            ID_24C1024 = 10
            ID_24C2048 = 11
            ID_24C4096 = 12
        End Enum

        ''' <summary>
        ''' Скорости интерфейса I2C / частота SCL.
        ''' </summary>
        Public Enum I2cSpeed
            ''' <summary>
            ''' Частота SCL 20 кГц.
            ''' </summary>
            LowSpeed = 0
            ''' <summary>
            ''' Частота SCL 100 кГц.
            ''' </summary>
            Standard = 1
            ''' <summary>
            ''' Частота SCL 400 кГц.
            ''' </summary>
            Fast = 2
            ''' <summary>
            ''' Частота SCL 750 кГц.
            ''' </summary>
            HighSpeed = 3
        End Enum

#End Region '/ENUMS

#Region "I2C"

        ''' <summary>
        ''' Задаёт частоту линии SCL / скорость передачи.
        ''' </summary>
        ''' <param name="speed">Скорость шины I2C.</param>
        Public Sub SetSpeed(speed As I2cSpeed)
            Dim res As Boolean = CH341SetStream(DeviceIndex, speed)
            If res Then
                Return
            End If
            Throw New Exception("Ошибка установки скорости шины I2C.")
        End Sub

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341StreamI2C(iIndex As Integer, iWriteLength As Integer, iWriteBuffer As Byte(), iReadLength As Integer, oReadBuffer As Byte()) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Универсальный метод чтения и/или записи.
        ''' </summary>
        ''' <param name="writeBuffer">Массив для записи. Первый байт это обычно I2C адрес, сдвинутый на 1 влево (например, если адрес 0x40, то первый элемент массива 0x80). Если NULL или пустой массив, то только чтение.</param>
        ''' <param name="readLength">Сколько байтов прочитать. Если 0, то только запись.</param>
        Public Function I2cReadWrite(writeBuffer As Byte(), readLength As Integer) As Byte()
            Dim wLen As Integer = 0
            If (writeBuffer IsNot Nothing) Then
                wLen = writeBuffer.Length
            End If
            Dim readBuffer(readLength - 1) As Byte
            Dim res As Boolean = CH341StreamI2C(DeviceIndex, wLen, writeBuffer, readLength, readBuffer)
            If res Then
                Return readBuffer
            End If
            Throw New Exception("Ошибка чтения/записи по I2C.")
        End Function

        ''' <summary>
        ''' Универсальный метод чтения и/или записи.
        ''' </summary>
        ''' <param name="address">I2C адрес ведомого.</param>
        ''' <param name="writeBuffer">Массив для записи. Если NULL или пустой массив, то только чтение.</param>
        ''' <param name="readLength">Сколько байтов прочитать. Если 0, то только запись.</param>
        Public Function I2cReadWrite(address As Integer, writeBuffer As Byte(), readLength As Integer) As Byte()
            Dim wLen As Integer = 0
            Dim newWriteBuf(0) As Byte
            If (writeBuffer IsNot Nothing) Then
                wLen = writeBuffer.Length + 1
                ReDim newWriteBuf(wLen - 1)
            End If
            newWriteBuf(0) = CByte(address << 1)
            Array.Copy(writeBuffer, 0, newWriteBuf, 1, writeBuffer.Length)
            Dim readBuffer(readLength - 1) As Byte
            Dim res As Boolean = CH341StreamI2C(DeviceIndex, wLen, newWriteBuf, readLength, readBuffer)
            If res Then
                Return readBuffer
            End If
            Throw New Exception("Ошибка чтения/записи по I2C.")
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341ReadI2C(iIndex As Integer, iDevice As Byte, iAddr As Byte, ByRef oByte As Byte) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Читает 1 байт по шине I2C.
        ''' </summary>
        ''' <param name="slaveAddress">Младшие 7 битов представляют i2c адрес ведомого устройства, старший бит - направление передачи (1 - чтение).</param>
        ''' <param name="register">Адрес регистра данных.</param>
        Public Function I2cRead(slaveAddress As Byte, register As Byte) As Byte
            slaveAddress = CByte(slaveAddress And &HFE)
            Dim b As Byte = 0
            Dim res As Boolean = CH341ReadI2C(DeviceIndex, slaveAddress, register, b)
            If res Then
                Return b
            End If
            Throw New Exception("Ошибка чтения по интерфейсу I2C.")
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341WriteI2C(iIndex As Integer, iDevice As Byte, iAddr As Byte, iByte As Byte) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Записывает 1 байт по шине I2C.
        ''' </summary>
        ''' <param name="slaveAddress">Младшие 7 битов представляют i2c адрес ведомого устройства, старший бит - направление передачи (0 - запись).</param>
        ''' <param name="register">Адрес регистра.</param>
        ''' <param name="dataToWrite">Байт для записи.</param>
        Public Sub I2cWrite(slaveAddress As Byte, register As Byte, dataToWrite As Byte)
            Dim res As Boolean = CH341WriteI2C(DeviceIndex, slaveAddress, register, dataToWrite)
            If res Then
                Return
            End If
            Throw New Exception("Ошибка записи по интерфейсу I2C.")
        End Sub

#End Region '/I2C

#Region "EEPROM via I2C"

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341ReadEEPROM(iIndex As Integer, iEepromID As EepromTypes, iAddr As Integer, iLength As Integer, oBuffer As Byte()) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Читает из I2C ПЗУ заданное число байтов на скорости примерно 56 кб/с.
        ''' </summary>
        ''' <param name="eepromType">Тип ПЗУ.</param>
        ''' <param name="address">Адрес чтения.</param>
        ''' <param name="length">Сколько байтов прочитать.</param>
        ''' <remarks>ПЗУ подключается по линиям: SDA - пин 23, SCL - пин 24.</remarks>
        Public Function EepromRead(eepromType As EepromTypes, address As Integer, length As Integer) As Byte()
            Dim readBuffer(length - 1) As Byte
            Dim res As Boolean = CH341ReadEEPROM(DeviceIndex, eepromType, address, length, readBuffer)
            If res Then
                Return readBuffer
            End If
            Throw New Exception("Невозможно прочитать данные из ПЗУ.")
        End Function

        <DllImport(DLL_PATH, SetLastError:=True, CallingConvention:=CallingConvention.StdCall)>
        Private Shared Function CH341WriteEEPROM(iIndex As Integer, iEepromID As EepromTypes, iAddr As Integer, iLength As Integer, iBuffer As Byte()) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function

        ''' <summary>
        ''' Записывает 
        ''' </summary>
        ''' <param name="eepromType">Тип ПЗУ.</param>
        ''' <param name="address">Адрес записи.</param>
        ''' <param name="writeBuffer">Данные для записи в ПЗУ.</param>
        Public Sub EepromWrite(eepromType As EepromTypes, address As Integer, writeBuffer As Byte())
            Dim res As Boolean = CH341WriteEEPROM(DeviceIndex, eepromType, address, writeBuffer.Length, writeBuffer)
            If res Then
                Return
            End If
            Throw New Exception("Невозможно записать данные в ПЗУ.")
        End Sub

#End Region '/EEPROM via I2C

    End Class '/I2C

End Namespace