# ch341_vb

This is the Visual Basic wrapper for **CH341DLL** library.
It realize all the available working modes of ch341a chip: `serial` (SPI, IIC) and `parallel` (EPP, MEM).

### Prerequisites

- USB driver for CH341 must be installed;
- ch341dll.dll must be placed near to your executable.

## Usage

First, imports Ch341 namespace:

```
Imports Ch341
```

### SPI example, VB.NET:

```
''' <summary>
''' Read ID register of BMP280 sensor.
''' </summary>
Private Sub CheckBmp280_spi()        
  Dim ch341 As New SpiMaster(0, True)
  Dim buf As Byte() = ch341.StreamSpi({&HD0}, 1)
  For Each b As Byte In buf
    Console.Write(b.ToString("X2"))
    Console.Write(" ")
  Next
  Console.WriteLine()
  ch341.CloseDevice()
End Sub
```
  
### I2C example, VB.NET:

```
Dim ch341 As New I2cMaster(0, True, I2cMaster.I2cSpeed.Standard)
```

## Additional documentation

More complex description:
- [SPI examples](https://soltau.ru/index.php/themes/dev/item/514-realizatsiya-interfejsa-spi-s-pomoshchyu-biblioteki-ch341dll-na-vb-net)
- [I2C examples](https://soltau.ru/index.php/themes/dev/item/512-realizatsiya-interfejsa-i2c-s-pomoshchyu-biblioteki-ch341dll-na-vb-net)
