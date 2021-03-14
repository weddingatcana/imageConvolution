Attribute VB_Name = "m_Functions"
Option Explicit

Public Enum greyType
    greyType_AVERAGE = 0
    greyType_LUMINANCE = 1
    greyType_DESATURATION = 2
    greyType_RED = 3
    greyType_GREEN = 4
    greyType_BLUE = 5
End Enum

Public Enum convKernel
    convKernel_SOBELX = 0
    convKernel_SOBELY = 1
    convKernel_GAUSSIAN3x3 = 2
    convKernel_GAUSSIAN5x5 = 3
    convKernel_BOXBLUR = 4
    convKernel_PREWITTX = 5
    convKernel_PREWITTY = 6
    convKernel_SHARPEN = 7
    convKernel_pseudoLAPLACIAN = 8
    convKernel_IDENTITY = 9
End Enum

Public Enum errorType
    'custom error handling
    'ex:
    'errorType_MODULE_FUNCTION_TYPE
    errorType_m_Functions_Convolve = 0
End Enum


Function Convolve(ByRef srcArray&(), ByRef convArray!(), Optional ByVal srcPadding As Boolean = True, Optional ByVal convLoop% = 1) As Long()

Rem Purpose: This was created to easily convolve general 2D arrays with a 2D square kernel. This function will be used in pixel routines _
             which is why all arrays are dimensioned for long integers. The optional boolean argument for padding is defaulted to be true. _
             This generates a field of zeros around our srcArray, such that the resultant convolved array doesnt shrink as a result of the _
             convolution. The number of rows and columns of zeros are calculated from the size of the kernel, convArray. _
 _
             I would prefer to pass the arrays to the function ByVal, but since I don't want them of type Variant, I am coerced by VBA to pass _
             them ByRef. Given that, as shown in the variable declarations, I have dimensioned copy arrays to load the information in and _
             manipulate as needed. {Derek Butler | 12/28/21 @ 9:14PM}
             
    
    Dim srcCopy&(), _
        srcCopyWithPadding&(), _
        convCopy!(), _
        outputArray&(), _
        srcMaxRow&, _
        srcMaxCol&, _
        convMaxRow&, _
        convMaxCol&, _
        outputMaxRow&, _
        outputMaxCol&, _
        srcPaddingNum&, _
        sum&, _
        base&, _
        offset&, _
        increment&, _
        i&, j&, p&, q&

        srcMaxRow = UBound(srcArray, 2)
        srcMaxCol = UBound(srcArray, 1)
        convMaxRow = UBound(convArray, 2)
        convMaxCol = UBound(convArray, 1)
        
        If convMaxRow <> convMaxCol Then
            'Custom error enumeration, like in the class module example online. "Odd Square Matrix Only"
        Else
            base = convMaxRow
            offset = base - 1
        End If
        
        ReDim convCopy(1 To convMaxCol, 1 To convMaxRow)
        
        'Our kernel is flipped both horizontally, and vertically. Thus true convolution; otherwise, it's a cross correlation.
        For i = 1 To convMaxRow
            For j = 1 To convMaxCol
            
                p = (i + offset) Mod base
                q = (j + offset) Mod base
                
                convCopy(i, j) = convArray(i + offset - 2 * p, j + offset - 2 * q)
                
            Next j
        Next i
        
        ReDim srcCopy(1 To srcMaxCol, 1 To srcMaxRow)
        srcCopy = srcArray
        
        If srcPadding Then
        
            srcPaddingNum = convMaxRow \ 2
            ReDim srcCopyWithPadding(1 To (srcMaxCol + 2 * srcPaddingNum), 1 To (srcMaxRow + 2 * srcPaddingNum))
            
            'Loop counter, that increments the repeated convolutions.
            increment = 1
            Do
            
                For j = 1 To srcMaxCol
                    For i = 1 To srcMaxRow
                    
                        'Need to offset indices.
                        srcCopyWithPadding(j + srcPaddingNum, i + srcPaddingNum) = srcCopy(j, i)
                        
                    Next i
                Next j
                
                ReDim outputArray(1 To srcMaxCol, 1 To srcMaxRow)
            
                For i = 1 To srcMaxRow
                    For j = 1 To srcMaxCol
                            
                        sum = 0
                        For p = 1 To convMaxRow
                            For q = 1 To convMaxCol
                                
                                'The moneyshot(tm)
                                sum = sum + srcCopyWithPadding(j + q - 1, i + p - 1) * convCopy(p, q)
                                
                            Next q
                        Next p
                        outputArray(j, i) = sum
                        
                    Next j
                Next i
                
                increment = increment + 1
                
                If increment > convLoop Then
                    Convolve = outputArray
                Else
                    srcCopy = outputArray
                End If
            
            Loop Until increment > convLoop
                  
        Else
                        
            'Loop counter, that increments the repeated convolutions.
            increment = 1
            Do
            
                srcMaxRow = UBound(srcCopy, 2)
                srcMaxCol = UBound(srcCopy, 1)
                
                outputMaxRow = srcMaxRow - convMaxRow + 1
                outputMaxCol = srcMaxCol - convMaxCol + 1
                ReDim outputArray(1 To outputMaxCol, 1 To outputMaxRow)
            
                'Bounds change to less than srcCopy when padding isnt involved.
                For i = 1 To outputMaxRow
                    For j = 1 To outputMaxCol
                            
                        sum = 0
                        For p = 1 To convMaxRow
                            For q = 1 To convMaxCol
                                
                                'The moneyshot(tm) v2.0
                                sum = sum + srcCopy(j + q - 1, i + p - 1) * convCopy(p, q)
                                
                            Next q
                        Next p
                        outputArray(j, i) = sum
                        
                    Next j
                Next i
            
                increment = increment + 1
                
                If increment > convLoop Then
                    Convolve = outputArray
                Else
                    ReDim srcCopy(1 To outputMaxCol, 1 To outputMaxRow)
                    srcCopy = outputArray
                End If
            
            Loop Until increment > convLoop
            
        End If
    
End Function

Function PreProcessedKernel(ByVal Kernel As convKernel) As Single()

Rem Purpose: This was created to easily select a convolution kernel without manually defining it before running the _
             Convolve function. Note that for both the Sobel and Prewitt kernels, the gradient magnitude must be calculated _
             from both the X and Y matrices before writing back to the window handle. Otherwise, you would only be defining _
             edges in either the X or Y direction and not both.


    Dim k!(), _
        f!, _
        Factor!
    
        If Kernel = convKernel_SOBELX Then
        
            ReDim k(1 To 3, 1 To 3)
            
            k(1, 1) = 1:    k(1, 2) = 0:    k(1, 3) = -1
            k(2, 1) = 2:    k(2, 2) = 0:    k(2, 3) = -2
            k(3, 1) = 1:    k(3, 2) = 0:    k(3, 3) = -1
            
            PreProcessedKernel = k
            
        ElseIf Kernel = convKernel_SOBELY Then
        
            ReDim k(1 To 3, 1 To 3)
            
            k(1, 1) = 1:    k(1, 2) = 2:    k(1, 3) = 1
            k(2, 1) = 0:    k(2, 2) = 0:    k(2, 3) = 0
            k(3, 1) = -1:   k(3, 2) = -2:   k(3, 3) = -1
            
            PreProcessedKernel = k
        
        ElseIf Kernel = convKernel_GAUSSIAN3x3 Then
        
            ReDim k(1 To 3, 1 To 3)
            Factor = 1 / 16
            f = Factor
            
            k(1, 1) = f * 1:    k(1, 2) = f * 2:    k(1, 3) = f * 1
            k(2, 1) = f * 2:    k(2, 2) = f * 4:    k(2, 3) = f * 2
            k(3, 1) = f * 1:    k(3, 2) = f * 2:    k(3, 3) = f * 1
            
            PreProcessedKernel = k

        ElseIf Kernel = convKernel_GAUSSIAN5x5 Then
        
            ReDim k(1 To 5, 1 To 5)
            Factor = 1 / 256
            f = Factor
            
            k(1, 1) = f * 1:    k(1, 2) = f * 4:    k(1, 3) = f * 6:    k(1, 4) = f * 4:    k(1, 5) = f * 1
            k(2, 1) = f * 4:    k(2, 2) = f * 16:   k(2, 3) = f * 24:   k(2, 4) = f * 16:   k(2, 5) = f * 4
            k(3, 1) = f * 6:    k(3, 2) = f * 24:   k(3, 3) = f * 36:   k(3, 4) = f * 24:   k(3, 5) = f * 6
            k(4, 1) = f * 4:    k(4, 2) = f * 16:   k(4, 3) = f * 24:   k(4, 4) = f * 16:   k(4, 5) = f * 4
            k(5, 1) = f * 1:    k(5, 2) = f * 4:    k(5, 3) = f * 6:    k(5, 4) = f * 4:    k(5, 5) = f * 1
            
            PreProcessedKernel = k
            
        ElseIf Kernel = convKernel_BOXBLUR Then
        
            ReDim k(1 To 3, 1 To 3)
            Factor = 1 / 9
            f = Factor
            
            k(1, 1) = f * 1:    k(1, 2) = f * 1:    k(1, 3) = f * 1
            k(2, 1) = f * 1:    k(2, 2) = f * 1:    k(2, 3) = f * 1
            k(3, 1) = f * 1:    k(3, 2) = f * 1:    k(3, 3) = f * 1
            
            PreProcessedKernel = k
            
        ElseIf Kernel = convKernel_PREWITTX Then
        
            ReDim k(1 To 3, 1 To 3)
            
            k(1, 1) = 1:    k(1, 2) = 0:    k(1, 3) = -1
            k(2, 1) = 1:    k(2, 2) = 0:    k(2, 3) = -1
            k(3, 1) = 1:    k(3, 2) = 0:    k(3, 3) = -1
            
            PreProcessedKernel = k
        
        ElseIf Kernel = convKernel_PREWITTY Then
        
            ReDim k(1 To 3, 1 To 3)
            
            k(1, 1) = 1:    k(1, 2) = 1:    k(1, 3) = 1
            k(2, 1) = 0:    k(2, 2) = 0:    k(2, 3) = 0
            k(3, 1) = -1:   k(3, 2) = -1:   k(3, 3) = -1
            
            PreProcessedKernel = k
            
        ElseIf Kernel = convKernel_SHARPEN Then
        
            ReDim k(1 To 3, 1 To 3)
            
            k(1, 1) = 0:    k(1, 2) = -1:   k(1, 3) = 0
            k(2, 1) = -1:   k(2, 2) = 5:    k(2, 3) = -1
            k(3, 1) = 0:    k(3, 2) = -1:   k(3, 3) = 0
            
            PreProcessedKernel = k
            
        ElseIf Kernel = convKernel_pseudoLAPLACIAN Then
        
            ReDim k(1 To 3, 1 To 3)
            
            k(1, 1) = -1:   k(1, 2) = -1:   k(1, 3) = -1
            k(2, 1) = -1:   k(2, 2) = 8:    k(2, 3) = -1
            k(3, 1) = -1:   k(3, 2) = -1:   k(3, 3) = -1
            
            PreProcessedKernel = k
            
        ElseIf Kernel = convKernel_IDENTITY Then
        
            ReDim k(1 To 3, 1 To 3)
            
            k(1, 1) = 0:    k(1, 2) = 0:    k(1, 3) = 0
            k(2, 1) = 0:    k(2, 2) = 1:    k(2, 3) = 0
            k(3, 1) = 0:    k(3, 2) = 0:    k(3, 3) = 0
            
            PreProcessedKernel = k
            
        End If
            
End Function

Function CVtoRGB(ByVal lngCV&) As Long()

Rem Purpose: The color value composed from all three color channels, red, green and blue are represented in base10 form _
             by the following equation: _
 _
             CV = (256^2 * R) + (256 * G) + (B) _
 _
             The decomposition of color value into the red, green and blue components are given by this function as an array per CV. _
             The range of color values is from 0 to 16,777,215. If outside of these bounds, the function returns the color black.
             

    Dim arrRGB&(), _
        lngBase&, _
        R&, G&, B&
    
        ReDim arrRGB(1 To 3)
        lngBase = 256
    
        If lngCV < (lngBase ^ 3) + 1 Then
    
            B = lngCV Mod lngBase
            G = (lngCV \ lngBase) Mod lngBase
            R = (lngCV \ lngBase ^ (2)) Mod lngBase
        
            arrRGB(1) = R
            arrRGB(2) = G
            arrRGB(3) = B
        
            CVtoRGB = arrRGB
        
        Else
    
            arrRGB(1) = 0
            arrRGB(2) = 0
            arrRGB(3) = 0
        
            CVtoRGB = arrRGB

        End If
    
End Function

Function RawDataToGrey(ByRef rawData() As Byte, Optional ByVal GreyscaleCalculation As greyType = greyType_AVERAGE) As Long()

Rem Purpose: Our data from GetBitmapBits comes to us in the byte array form: rawData(color,y,x). Since it comes in ByRef, we transfer it into a copy _
             array and convert it from byte to long. Again, this helps with calculation speeds. From there, we convert the RGB values to greyscale _
             following the greyscale calculation procedure outlined in the enumeration. This function then returns a 2D long array of grey values.


    Dim rawCopy&(), _
        rawMaxRow&, _
        rawMaxCol&, _
        rawMaxColor&, _
        Grey&(), _
        R&, G&, B&, _
        i&, j&, k&
        
        rawMaxRow = UBound(rawData, 3)
        rawMaxCol = UBound(rawData, 2)
        rawMaxColor = UBound(rawData, 1)
        
        ReDim rawCopy(1 To rawMaxColor, 1 To rawMaxCol, 1 To rawMaxRow)
        
        'Need to convert array from type byte to long. A CLng() wrapper wont due when applied to whole array. _
        We will need to apply to all elements. UGH. I only do this because VBA is faster when using longs.
        
        For k = 1 To rawMaxColor
            For j = 1 To rawMaxCol
                For i = 1 To rawMaxRow
                
                    rawCopy(k, j, i) = CLng(rawData(k, j, i))
                
                Next i
            Next j
        Next k

        ReDim Grey(1 To rawMaxCol, 1 To rawMaxRow)
        
        For j = 1 To rawMaxCol
            For i = 1 To rawMaxRow
            
                R = rawCopy(1, j, i)
                G = rawCopy(2, j, i)
                B = rawCopy(3, j, i)
                'Alpha = rawCopy(4, j, i)
                
                Grey(j, i) = Greyscale(R, G, B, GreyscaleCalculation)
                
            Next i
        Next j
        
        RawDataToGrey = Grey

End Function

Function Greyscale&(ByVal R&, ByVal G&, ByVal B&, Optional ByVal GreyscaleCalculation As greyType = greyType_AVERAGE)

Rem Purpose: Greyscale comes in many forms, and I thought it would be interesting to display the different ways it can be _
             calculated. Usually, the easiest and most intuitive way would be to average the RGB values, which is why I set _
             it to be the default if no option is chosen. Photoshop, GIMP and other programs try to correct the grey values _
             to what the eye is most sensitive to, the green color channels. This is evidenced by the weight given to the _
             green component in the greyType_Luminance calculation. Subtle changes, but noticeable to the eye.


    Dim k&(), _
        max&, _
        min&, _
        temp&, _
        i&, j&
        
        If GreyscaleCalculation = greyType_AVERAGE Then
        
            Greyscale = (R + G + B) / 3
             
        ElseIf GreyscaleCalculation = greyType_LUMINANCE Then
        
            Greyscale = R * 0.3 + G * 0.59 + B * 0.11
            
        ElseIf GreyscaleCalculation = greyType_DESATURATION Then
        
            ReDim k(1 To 3)
            k(1) = R: k(2) = G: k(3) = B
           
           ' Sorts array from largest to smallest. Quick & dirty bubblesort.
            For i = LBound(k) To UBound(k)
                For j = i + 1 To UBound(k)
                   If k(i) < k(j) Then
                    
                        temp = k(j)
                        k(j) = k(i)
                        k(i) = temp
                        
                    End If
                Next j
            Next i
            
            max = k(1)
            min = k(3)
            Greyscale = (max + min) / 2
            
        ElseIf GreyscaleCalculation = greyType_RED Then
        
            Greyscale = R
            
        ElseIf GreyscaleCalculation = greyType_GREEN Then
        
            Greyscale = G
            
        ElseIf GreyscaleCalculation = greyType_BLUE Then
        
            Greyscale = B
            
        End If

End Function

Function Dimension2Dto3D(ByRef src2D&()) As Byte()

Rem Purpose: By the time we reach this function, we have taken a 3D byte array of color and turned it into a 2D long array of grey. The _
             conversion was necessary to make processing quicker. Now, after convolutions with kernels, or just graying the picture _
             with the desired calculation we then come to the task of displaying it back to the window handle. That's where this function _
             comes into play. It is important to note that while the API for both Get/SetBitmapBits refer to a pointer to a buffer _
             to store the image data as declared for anything greater than a byte (0-255), I tried continuing to use long arrays and the result was _
             a messed up image, with seemingly only the blue channel displayed. _
 _
             As such, this function will convert our long arrays to byte arrays. This newly created byte array will copy over the grey data _
             from our long array to every color channel except alpha. So there is a check to see if you are on a 32bit color system, and _
             if so to disregard the alpha entirely during the copy procedure. There is a small check for the color architecture, as I didn't want to _
             hard code it into the system as 3 or 4. This dynamic procedure is required now, because before the data given out by our GetObject/GetBitmapBits API _
             already did that work behind the scenes for us. We merely referenced the bounds as it was an input to our functions prior. We _
             cant access that data directly now, so we use this method.
             

    Const BitsPixel& = &HC '12
    Const hWnd As LongPtr = &H0 '0
    Const base& = 256

    Dim src3D() As Byte, _
        srcMaxRow&, _
        srcMaxCol&, _
        srcMaxColor&, _
        srcMaxColorCorrection&, _
        Check&, _
        hDC As LongPtr, _
        i&, j&, k&

        'Obtain handle to desktop window, find if we're on a 24bit/32bit color architecture.
        hDC = GetDC(hWnd)
        srcMaxColor = GetDeviceCaps(hDC, BitsPixel) \ 8
        Check = ReleaseDC(hWnd, hDC)
        
        If Check = 1 Then
            'released
        ElseIf Check = 0 Then
            'not released
        End If
        
        srcMaxRow = UBound(src2D, 2)
        srcMaxCol = UBound(src2D, 1)
        
        'We still want to include the alpha channel in our array, if on a 32bit system. Hence, dimensioning to srcMaxColor. _
        SetBitmapBits will need that extra alpha channel if on a 32bit system.
        ReDim src3D(1 To srcMaxColor, 1 To srcMaxCol, 1 To srcMaxRow)
        
        'But, as following from the comment above, we don't want to copy src2D data into the alpha channel (#4).
        If srcMaxColor > 3 Then
            srcMaxColorCorrection = 3
        Else
            srcMaxColorCorrection = srcMaxColor
        End If
        
        For k = 1 To srcMaxColorCorrection
            For j = 1 To srcMaxCol
                For i = 1 To srcMaxRow
                
                    src3D(k, j, i) = CByte(Abs(src2D(j, i) Mod base))
                                
                Next i
            Next j
        Next k
        
        Dimension2Dto3D = src3D

End Function

Function Gradient(ByRef arr1&(), ByRef arr2&()) As Long()

Rem Purpose: For some kernel filters, we pass unique symmetrical operators to get the response in the X or Y direction. These directional derivatives can _
             show results by themselves but when the magnitude of both are taken, we see the full effect. This function will take two 2D long arrays and _
             convert them into a singular 2D long array that represents the magnitude of both inputs.


    Dim arrGradient&(), _
        arr1MaxRow&, _
        arr1MaxCol&, _
        arr2MaxRow&, _
        arr2MaxCol&, _
        i&, j&, _
        x&, y&
        
        arr1MaxRow = UBound(arr1, 2)
        arr1MaxCol = UBound(arr1, 1)
        arr2MaxRow = UBound(arr2, 2)
        arr2MaxCol = UBound(arr2, 1)
        
        'Check if both arrays are the same size.
        If arr1MaxRow = arr2MaxRow And arr1MaxCol = arr2MaxCol Then
        
            ReDim arrGradient(1 To arr1MaxCol, 1 To arr1MaxRow)
            
            For j = 1 To arr1MaxCol
                For i = 1 To arr1MaxRow
                
                    'Theta = Atn(Gy/Gx)
                    x = arr1(j, i) * arr1(j, i)
                    y = arr2(j, i) * arr2(j, i)
                    arrGradient(j, i) = Sqr(x + y)
                
                Next i
            Next j
        
            Gradient = arrGradient
        
        Else
        End If

End Function

Function LinearInterpolation(ByRef rawData&(), ByVal srcWidth, ByVal srcHeight) As Long()

    Dim arrCopy&(), _
        arrRescaled&(), _
        rawMaxRow&, _
        rawMaxCol&, _
        strLimit&, _
        aspectRatio!, _
        newWidth&, _
        newHeight&, _
        x_j&, y_i&, _
        i&, j&
        
        strLimit = 1024
        aspectRatio = srcWidth / srcHeight
        rawMaxRow = UBound(rawData, 2)
        rawMaxCol = UBound(rawData, 1)
        ReDim arrCopy(1 To rawMaxCol, 1 To rawMaxRow)
        
        If srcWidth <= strLimit And srcWidth > 0 Then
            
            arrCopy = rawData
            LinearInterpolation = arrCopy
            Exit Function
            
        ElseIf srcWidth > strLimit Then
        
            newWidth = strLimit
            newHeight = newWidth / aspectRatio
            
            arrCopy = rawData
            ReDim arrRescaled(1 To newWidth, 1 To newHeight)
        
            For j = 1 To newWidth
                For i = 1 To newHeight
                
                    x_j = Round((j / newWidth) * rawMaxCol)
                    y_i = Round((i / newHeight) * rawMaxRow)
                    arrRescaled(j, i) = arrCopy(x_j, y_i)
                
                Next i
            Next j
            
            LinearInterpolation = arrRescaled
        
        End If

End Function

Function ExportAsASCII(ByRef rawData&()) As String()

    Dim arrASCII$(), _
        arrFinal$(), _
        arrCopy&(), _
        strASCII$, _
        rawMaxRow&, _
        rawMaxCol&, _
        max&, _
        interval&, _
        strLength&, _
        i&, j&, k&
        
        strLength = 1
        'strASCII = "$@B%8&WM#0*oahkbdpqwmZOLCJYXzcuxrjft/\|()1{}[]?-_+~<>i!lI;:,^`'."
        'strASCII = "$%&#*kmOXx|?~!:."
        strASCII = "@%#*+=-."
        'strASCII = "@0Ox+!:."
        ReDim arrASCII(1 To Len(strASCII))
        
        For i = 1 To Len(strASCII)
            arrASCII(i) = Mid$(strASCII, i, strLength)
        Next i
        
        rawMaxRow = UBound(rawData, 2)
        rawMaxCol = UBound(rawData, 1)
        ReDim arrCopy(1 To rawMaxCol, 1 To rawMaxRow)
        ReDim arrFinal(1 To rawMaxCol, 1 To rawMaxRow)
        
        arrCopy = rawData
        max = 256
        interval = max \ Len(strASCII)
        
        For j = 1 To rawMaxCol
            For i = 1 To rawMaxRow
            
                k = (max - arrCopy(j, i)) \ interval
                If k = 0 Then
                    k = k + 1
                End If
                arrFinal(j, i) = arrASCII(k)
            
            Next i
        Next j

        ExportAsASCII = arrFinal

End Function

Function ExportASCIIPicture(ByRef rawASCII$(), Optional TextFile As Boolean = True, Optional HTMLFile As Boolean = False) As Boolean
    
    Dim wkbkASCII As Workbook, _
        wkshtASCII As Worksheet, _
        rngASCII As Range, _
        strFilePath, _
        arrCopy$(), _
        rawMaxRow&, _
        rawMaxCol, _
        boolCheck As Boolean, _
        i&, j&
        
        rawMaxRow = UBound(rawASCII, 2)
        rawMaxCol = UBound(rawASCII, 1)
        ReDim arrCopy(1 To rawMaxRow, 1 To rawMaxCol)
        
        For j = 1 To rawMaxCol
            For i = 1 To rawMaxRow
                arrCopy(i, j) = rawASCII(j, i)
            Next i
        Next j
        
        boolCheck = TempFolderExists
        
        If boolCheck = True Then
        
            If (HTMLFile = True And TextFile = True) Or (HTMLFile = False And TextFile = False) Then
        
            ElseIf HTMLFile = True Then
                
                Application.DisplayAlerts = False
                Set wkbkASCII = Application.Workbooks.Add
                Set wkshtASCII = wkbkASCII.Worksheets.Add
                
                With wkshtASCII
                    .Name = "ASCII"
                    .Columns.ColumnWidth = 1
                    .Rows.RowHeight = 10
                    .Activate
                End With
                
                With wkbkASCII
                    .Worksheets("Sheet1").Delete
                End With
                
                ActiveWindow.DisplayGridlines = False
                
                Set rngASCII = wkshtASCII.Range(Cells(1, 1), Cells(rawMaxRow, rawMaxCol))
                rngASCII = arrCopy
                
                strFilePath = "C:\temp\ASCII.html"
                With wkbkASCII
                    .SaveAs strFilePath, xlHtml
                    .Close
                End With
                
                Application.DisplayAlerts = True
                Set wkbkASCII = Nothing
                Set wkshtASCII = Nothing
                ExportASCIIPicture = True
                
            ElseIf TextFile = True Then
            
                strFilePath = "C:\temp\ASCII.txt"
                boolCheck = WriteToTextFile(arrCopy, strFilePath)
                ExportASCIIPicture = True
            
            End If
            
        End If

End Function

Function WriteToTextFile(ByRef arrASCII$(), ByVal strFilePath$) As Boolean

    Dim FSO As Object, _
        txtFile As Object, _
        arrCopy$(), _
        arrMaxRow&, _
        arrMaxCol&, _
        concatString$, _
        i&, j&
        
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set txtFile = FSO.CreateTextFile(strFilePath)
        
        arrMaxRow = UBound(arrASCII, 1)
        arrMaxCol = UBound(arrASCII, 2)
        ReDim arrCopy(1 To arrMaxRow, 1 To arrMaxCol)
        
        arrCopy = arrASCII
        
        For i = 1 To arrMaxRow
            For j = 1 To arrMaxCol
            
                concatString = concatString & arrCopy(i, j)
                
            Next j
            
            txtFile.Write concatString & vbCrLf
            concatString = ""
            
        Next i
        
        txtFile.Close
        WriteToTextFile = True
        Set FSO = Nothing
        Set txtFile = Nothing

End Function

Function TempFolderExists() As Boolean

    Dim FSO As Object, _
        strFolder$
        
        Set FSO = CreateObject("Scripting.FileSystemObject")
        strFolder = "C:\temp"
        
        If FSO.FolderExists(strFolder) Then
            TempFolderExists = True
        Else
            FSO.CreateFolder (strFolder)
            TempFolderExists = True
        End If
        
        Set FSO = Nothing

End Function

Function ExportAsCSV&(ByRef rawData&())

End Function

Function GenerateDeviceIndependentBitmapGDIPlus(ByVal ImageFilePath$) As LongPtr

    #If Win64 Then
    
        Dim inputGDI As GdiplusStartupInput, _
            tokenGDI As LongPtr, _
            bmpGDI As LongPtr, _
            hBmpGDI As LongPtr

    #Else
    
        Dim inputGDI As GdiplusStartupInput, _
            tokenGDI&, _
            bmpGDI&, _
            hBmpGDI&
    
    #End If
    
            inputGDI.GdiplusVersion = 1
            
            If GdiplusStartup(tokenGDI, inputGDI) = 0 Then
                If GdipCreateBitmapFromFile(StrPtr(ImageFilePath), bmpGDI) = 0 Then
                    GdipCreateHBITMAPFromBitmap bmpGDI, hBmpGDI, 0
                    GdipDisposeImage bmpGDI
                End If
            End If
            
            GdiplusShutdown tokenGDI
            GenerateDeviceIndependentBitmapGDIPlus = hBmpGDI
    
End Function

Function ChooseImageFile$()

    Dim FDO As FileDialog, _
        SelectionChosen&
    
        Set FDO = Application.FileDialog(msoFileDialogFilePicker)
        SelectionChosen = -1
        
        With FDO
            .InitialFileName = "C:\"
            .Title = "Choose Image File"
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Allowed Image Extensions", "*.bmp; *.gif; *.jpg; *.jpeg; *.png; *.tiff; *.tif; *.dib; *.wmf; *.emf"
            
            If .Show = SelectionChosen Then
                ChooseImageFile = .SelectedItems(1)
            Else
            End If
            
        End With
    
        Set FDO = Nothing

End Function

Function ChooseImageFileExtension$(ByVal SelectedItem$)

    Dim strExtension$, _
        strDelimiter$, _
        DelimiterPosition&, _
        MaxExtensionLength&

        MaxExtensionLength = 5
        strDelimiter = "."
        strExtension = SelectedItem
        
        DelimiterPosition = InStr(Len(strExtension) - MaxExtensionLength, strExtension, strDelimiter)
        
        strExtension = Right(strExtension, Len(strExtension) - DelimiterPosition + 1)
        ChooseImageFileExtension = strExtension

End Function

Function DisplayImage(ByRef Picture As typePic, ByRef RefIID As GUID, ByRef objPic As IPicture, Optional ByVal strOutputImage$ = "", Optional ByVal strExtension$ = "")

    Dim Check&, _
        CompletionStatus&
        
        CompletionStatus = 1
        Check = m_WIN32API.OleCreatePictureIndirect(Picture, RefIID, CompletionStatus, objPic)
        
        If strOutputImage <> "" And strExtension <> "" Then
            stdole.SavePicture objPic, strOutputImage & strExtension
        End If
        
        frmPixel.frmPicture.Picture = objPic
        Set objPic = Nothing
    
End Function

Function SourceImageDimensions(ByVal strInputImage$) As Long()

    Dim arrOutput&(), _
        WIA As Object
        
        Set WIA = CreateObject("WIA.ImageFile")
        ReDim arrOutput(1 To 2)

        If WIA Is Nothing Then
            Exit Function
        End If

        WIA.LoadFile strInputImage
        arrOutput(1) = WIA.Width
        arrOutput(2) = WIA.Height
        
        Set WIA = Nothing

        SourceImageDimensions = arrOutput

End Function

Function GenerateDeviceDependentBitmap(ByVal srcWidth&, ByVal srcHeight&) As LongPtr

    Const srcCopy = &HCC0020
    Const TwipToPixel! = 1.333333

    #If Win64 Then

        Dim hWnd As LongPtr, _
            hBmp As LongPtr, _
            hBmpPrev As LongPtr, _
            hDC As LongPtr, _
            hDCMem As LongPtr, _
            Check&, _
            frmWidth&, _
            frmHeight&

    #Else
    
        Dim hWnd&, _
            hBmp&, _
            hBmpPrev&, _
            hDC&, _
            hDCMem&, _
            Check&, _
            frmWidth&, _
            frmHeight&
            
    #End If
            
            With frmPixel.frmPicture
                frmWidth = .Width * TwipToPixel
                frmHeight = .Height * TwipToPixel
                hWnd = .[_GethWnd]
            End With
            
            hDC = GetDC(hWnd)
            hDCMem = CreateCompatibleDC(hDC)
            hBmp = CreateCompatibleBitmap(hDC, srcWidth, srcHeight)
            hBmpPrev = SelectObject(hDCMem, hBmp)
            Check = StretchBlt(hDCMem, 0, 0, srcWidth, srcHeight, hDC, 0, 0, frmWidth, frmHeight, srcCopy)
            hBmp = SelectObject(hDCMem, hBmpPrev)
            
            GenerateDeviceDependentBitmap = hBmp

End Function

