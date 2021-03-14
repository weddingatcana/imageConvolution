Attribute VB_Name = "m_Main"
Option Explicit

    Public globalImageDataGrey&()
    Public globalImageDataRGB() As Byte
    
    #If Win64 Then
    
        Public Type GdiplusStartupInput
            GdiplusVersion              As Long
            DebugEventCallback          As LongPtr
            SuppressBackgroundThread    As Long
            SuppressExternalCodecs      As Long
        End Type
        
        Public Type Bitmap
            Type                        As Long
            Width                       As Long
            Height                      As Long
            WidthBytes                  As Long
            Planes                      As Integer
            BitsPixel                   As Integer
            Bits                        As LongPtr
        End Type
        
        Public Type GUID
            Data1                       As Long
            Data2                       As Integer
            Data3                       As Integer
            Data4(7)                    As Byte
        End Type
        
        Public Type typePic
            Size                        As Long
            Type                        As Long
            hBmp                        As LongPtr
            hPal                        As LongPtr
        End Type

    #Else
        
        Public Type GdiplusStartupInput
            GdiplusVersion              As Long
            DebugEventCallback          As Long
            SuppressBackgroundThread    As Long
            SuppressExternalCodecs      As Long
        End Type
        
        Public Type Bitmap
            Type                        As Long
            Width                       As Long
            Height                      As Long
            WidthBytes                  As Long
            Planes                      As Integer
            BitsPixel                   As Integer
            Bits                        As Long
        End Type
        
        Public Type GUID
            Data1                       As Long
            Data2                       As Integer
            Data3                       As Integer
            Data4(7)                    As Byte
        End Type
        
        Public Type typePic
            Size                        As Long
            Type                        As Long
            hBmp                        As Long
            hPal                        As Long
        End Type
    #End If

Public Sub Main()

    Const CF_BITMAP& = &H1

    #If Win64 Then

        Dim hBmp As LongPtr, _
            srcWidth&, _
            srcHeight&, _
 _
            Picture As typePic, _
            objPic As IPicture, _
            RefIID As GUID, _
            structBMP As Bitmap, _
 _
            strInputImage$, _
            strExtension$, _
            strOutputImage$, _
            arrImageDimensions&()

    #Else
    
        Dim hBmp&, _
            srcWidth&, _
            srcHeight&, _
 _
            Picture As typePic, _
            objPic As IPicture, _
            RefIID As GUID, _
            structBMP As Bitmap, _
 _
            strInputImage$, _
            strExtension$, _
            strOutputImage$, _
            arrImageDimensions&()
  
    #End If
    
            With RefIID
                .Data1 = &H7BF80980
                .Data2 = &HBF32
                .Data3 = &H101A
                .Data4(0) = &H8B
                .Data4(1) = &HBB
                .Data4(2) = &H0
                .Data4(3) = &HAA
                .Data4(4) = &H0
                .Data4(5) = &H30
                .Data4(6) = &HC
                .Data4(7) = &HAB
            End With
            
            strInputImage = ChooseImageFile
            strExtension = ChooseImageFileExtension(strInputImage)
            strOutputImage = "C:\temp\test"
            
            ReDim arrImageDimensions(1 To 2)
            arrImageDimensions = SourceImageDimensions(strInputImage)
            srcWidth = arrImageDimensions(1)
            srcHeight = arrImageDimensions(2)
            
            hBmp = GenerateDeviceIndependentBitmapGDIPlus(strInputImage)
        
            With Picture
                .Size = LenB(Picture)
                .Type = CF_BITMAP
                .hBmp = hBmp
                .hPal = 0&
            End With
        
            DisplayImage Picture, RefIID, objPic
            DoEvents
            hBmp = GenerateDeviceDependentBitmap(srcWidth, srcHeight)

            GetObject hBmp, LenB(structBMP), structBMP
            
            ReDim globalImageDataRGB(1 To (structBMP.BitsPixel \ 8), 1 To structBMP.Width, 1 To structBMP.Height)
            GetBitmapBits hBmp, structBMP.WidthBytes * structBMP.Height, globalImageDataRGB(1, 1, 1)
            
            ReDim globalImageDataGrey(1 To srcWidth, 1 To srcHeight)
            globalImageDataGrey = RawDataToGrey(globalImageDataRGB, greyType_LUMINANCE)
            globalImageDataRGB = Dimension2Dto3D(globalImageDataGrey)
            SetBitmapBits hBmp, structBMP.WidthBytes * structBMP.Height, globalImageDataRGB(1, 1, 1)
            
            Dim s$()
            ReDim s(1 To srcWidth, 1 To srcHeight)
            
            s = ExportAsASCII(LinearInterpolation(globalImageDataGrey, srcWidth, srcHeight))
            ExportASCIIPicture s, True
    
            With Picture
                .Size = LenB(Picture)
                .Type = CF_BITMAP
                .hBmp = hBmp
                .hPal = 0&
            End With
    
            DisplayImage Picture, RefIID, objPic, strOutputImage, strExtension
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
  
    Exit Sub

Errors:
  
    MsgBox Err.Description
    
End Sub
