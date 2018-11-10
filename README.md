# QRCodeLibVBA
QRCodeLibVBA is a QR code generation library written in Excel VBA.
Generate model 2 code symbols based on JIS X 0510.

## Feature
- It corresponds to numbers, alphanumeric characters, 8 bit bytes, and kanji mode
- Can create split QR code
- Can be saved to 1 bpp or 24 bpp BMP file (DIB)
- It can be obtained as 1 bpp or 24 bpp IPicture object
- Image coloring (foreground color / background color) can be specified
- Character code in 8 bit byte mode can be specified


## Quick start
Please refer to QRCodeLib.xlam in 32bit version Excel.


## How to use
### Example 1. Indicates the minimum code of the QR code consisting of a single symbol (not a split QR code).

```vbnet
Public Sub Example()
    Dim sbls As Symbols
    Set sbls = CreateSymbols()
    sbls.AppendString "012345abcdefg"

    Dim pict As stdole.IPicture
    Set pict = sbls(0).Get24bppImage()
    
End Sub
```

### Example 2. Specify the error correction level
Create a Symbols object by setting the value of the ErrorCorrectionLevel enumeration to the argument of the CreateSymbols function.

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(ErrorCorrectionLevel.H)
```

### Example 3. Specify upper limit of model number
Create a Symbols object by setting arguments of the CreateSymbols function.
```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(maxVer:=10)
```

### Example 4. Specify the character code to use in 8-bit byte mode
Create a Symbols object by setting arguments of the CreateSymbols function.
```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(byteModeCharsetName:="utf-8")
```

### Example 5. Create divided QR code
Create a Symbols object by setting arguments of the CreateSymbols function. If you do not specify the upper limit of the model number, it will be split up to model number 40 as the upper limit.
```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(allowStructuredAppend:=True)
```

An example of dividing when exceeding model number 1 and acquiring IP picture object of each QR code is shown below.

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols(maxVer:=1, allowStructuredAppend:=True)
sbls.AppendString "abcdefghijklmnopqrstuvwxyz"

Dim pict As stdole.IPicture
Dim sbl As Symbol

For Each sbl In sbls
    Set pict = sbl.Get24bppImage()
Next
```

### Example 6. Save to BMP file
Use the Save1bppDIB, or Save 24bppDIB method of the Symbol class.

```vbnet
Dim sbls As Symbols
Set sbls = CreateSymbols()
sbls.AppendString "012345abcdefg"

sbls(0).Save1bppDIB "D:\qrcode1bpp1.bmp"
sbls(0).Save1bppDIB "D:\qrcode1bpp2.bmp", 10 ' 10 pixels per module
sbls(0).Save24bppDIB "D:\qrcode24bpp1.bmp"
sbls(0).Save24bppDIB "D:\qrcode24bpp2.bmp", 10 ' 10 pixels per module
```
