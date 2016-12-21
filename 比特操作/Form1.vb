Imports System.Runtime.InteropServices

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '创建 0~255 的 Byte 数组
        Dim BitsLength As Integer = Marshal.SizeOf(Of Byte) * 8
        Dim ByteList(2 ^ BitsLength - 1) As Byte
        Debug.Print("生成长度为 {0} 的 Byte() 数组.", ByteList.Length)

        For Index As Integer = LBound(ByteList) To UBound(ByteList)
            ByteList(Index) = Index
        Next
        Debug.Print("初始化 Byte() 数组.")

        '把 Byte 数组转换为比特数组
        Dim BitArray As BitArray = New BitArray(ByteList)
        Debug.Print("从 Byte() 数组创建比特数组.")

        Debug.Print("比特数组托管内存地址：" & GCHandle.ToIntPtr(GCHandle.Alloc(BitArray)).ToString())

        '输出比特数组
        Debug.Print("开始输出比特数组.")
        PrintBitArray(BitArray, BitsLength)
        Debug.Print("输出比特数组完毕.")

        '按 Byte 输出比特数组
        Debug.Print("按 Byte 输出比特数组.")
        For Index As Integer = 0 To BitArray.Length - 1 Step BitsLength
            Debug.Print(Index.ToString("0000") & " : " & GetDecimalNumber(BitArray, Index, BitsLength))
        Next

        '按 Int16 输出比特数组
        Debug.Print("按 Short 输出比特数组.")
        BitsLength = Marshal.SizeOf(Of Short) * 8
        For Index As Integer = 0 To BitArray.Length - 1 Step BitsLength
            Debug.Print(Index.ToString("0000") & " : " & GetDecimalNumber(BitArray, Index, BitsLength))
        Next

        '按 Integer 输出比特数组
        Debug.Print("按 Integer 输出比特数组.")
        BitsLength = Marshal.SizeOf(Of Integer) * 8
        For Index As Integer = 0 To BitArray.Length - 1 Step BitsLength
            Debug.Print(Index.ToString("0000") & " : " & GetDecimalNumber(BitArray, Index, BitsLength))
        Next

        '按 Long 输出比特数组
        Debug.Print("按 Long 输出比特数组.")
        BitsLength = Marshal.SizeOf(Of Long) * 8
        For Index As Integer = 0 To BitArray.Length - 1 Step BitsLength
            Debug.Print(Index.ToString("0000") & " : " & GetDecimalNumber(BitArray, Index, BitsLength))
        Next

        '反转比特数组所有位
        Debug.Print("反转比特数组所有位.")
        BitArray.Not()
        PrintBitArray(BitArray, BitsLength)

        '重置比特数组里所有位
        Debug.Print("重置比特数组里所有位.")
        BitArray.SetAll(False)
        BitsLength = Marshal.SizeOf(Of Integer) * 8
        For Index As Integer = 0 To 31
            BitArray.Set(Index + Index * 32, True)
            BitArray.Set(BitArray.Length - Index * 33 - 1, True)
        Next
        PrintBitArray(BitArray, BitsLength)

        '使用 BitConverter 转换
        Debug.Print("按 Short 输出比特数组.")
        BitsLength = Marshal.SizeOf(Of Short) * 8
        For Index As Integer = 0 To ByteList.Length - 1 Step Marshal.SizeOf(Of Short) / Marshal.SizeOf(Of Byte)
            Debug.Print(Index.ToString("0000") & " : " & BitConverter.ToInt16(ByteList, Index))
        Next

        Debug.Print("按 Integer 输出比特数组.")
        BitsLength = Marshal.SizeOf(Of Integer) * 8
        For Index As Integer = 0 To ByteList.Length - 1 Step Marshal.SizeOf(Of Integer) / Marshal.SizeOf(Of Byte)
            Debug.Print(Index.ToString("0000") & " : " & BitConverter.ToInt32(ByteList, Index))
        Next

        Debug.Print("按 Long 输出比特数组.")
        BitsLength = Marshal.SizeOf(Of Long) * 8
        For Index As Integer = 0 To ByteList.Length - 1 Step Marshal.SizeOf(Of Long) / Marshal.SizeOf(Of Byte)
            Debug.Print(Index.ToString("0000") & " : " & BitConverter.ToInt64(ByteList, Index))
        Next

        Debug.Print("按 Double 输出比特数组.")
        BitsLength = Marshal.SizeOf(Of Double) * 8
        For Index As Integer = 0 To ByteList.Length - 1 Step Marshal.SizeOf(Of Double) / Marshal.SizeOf(Of Byte)
            Debug.Print(Index.ToString("0000") & " : " & BitConverter.ToDouble(ByteList, Index))
        Next
    End Sub

    ''' <summary>
    ''' 输出比特数组
    ''' </summary>
    ''' <param name="BitArray">比特数组</param>
    ''' <param name="SplitLength">换行分割长度</param>
    Private Sub PrintBitArray(BitArray As BitArray, SplitLength As Integer)
        Dim BitString As String = vbNullString
        Dim ChildBitString As String = vbNullString

        For Index As Integer = 0 To BitArray.Length - 1
            If (Index Mod SplitLength) = 0 Then
                BitString &= ChildBitString & vbCrLf
                ChildBitString = vbNullString
            End If
            ChildBitString = IIf(BitArray.Get(Index), "1", "0") & ChildBitString
        Next

        Debug.Print(BitString)
    End Sub

    ''' <summary>
    ''' 返回从比特数组指定位置读取指定长度的部分转换为的十进制数字
    ''' </summary>
    ''' <param name="BitArray">比特数组</param>
    ''' <param name="StartIndex">开始位置</param>
    ''' <param name="Length">读取长度</param>
    ''' <returns></returns>
    Private Function GetDecimalNumber(BitArray As BitArray, StartIndex As Integer, Length As Integer) As Object
        Dim RObject As Object = 0
        For Index As Integer = 0 To Length - 1
            If BitArray.Item(StartIndex + Index) Then RObject += 1 << Index
        Next
        Return RObject
    End Function

End Class
