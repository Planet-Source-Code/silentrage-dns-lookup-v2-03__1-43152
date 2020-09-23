Attribute VB_Name = "modDNS"
Public Type DNSHeader   '12 bytes
    ID As Long            '2 Bytes
    'Query/Response
      'True = Response
      'False = Query
    QR As Boolean         '1 bit
    'Operation Code
      '0 = Standard Query
      '1 = Inverse Query
      '2 = Server Status Request
      '3-15 = Reserved
    OpCode As Byte        '4 bits
    'Authoritive Answer
    AA As Boolean         '1 bit
    'Truncated
    TC As Boolean         '1 bit
    'Recursion Desired
    RD As Boolean         '1 bit
    'Recursion Available
    RA As Boolean         '1 bit
    'Reserved
    Z As Byte             '3 bits
    'Response Code
      '0 = No Error
      '1 = Format Error - The name server was unable to interpret the query
      '2 = Server Failure - The name server was unable to process this query due to a problem with the name server
      '3 = Name Error - Meaningful only for responses from an authoritative name server, this code signifies that the domain name referenced in the query does not exist.
      '4 = Not Implemented - The name server does not support the requested kind of query
      '5 = Refused - The name server refuses to perform the specified operation for policy reasons.
      '6-15 = Reserved
    RCode As Byte         '4 bits
    'Question Record Count
    QDCount As Long       '2 Bytes
    'Answer Record Count
    ANCount As Long       '2 Bytes
    'Authoritative Name Server Count
    NSCount As Long       '2 Bytes
    'Additional Record Count
    ARCount As Long       '2 Bytes
End Type
Public Type DNSQuestion
    QName As String
    QType As Long         '2 Bytes
    QClass As Long        '2 Bytes
End Type
Public Type DNSRecord
    RName As String
    RType As Long         '2 Bytes
    RClass As Long        '2 Bytes
    TTL As Double         '4 Bytes
    RDLength As Long      '2 Bytes
    RData As Variant
End Type
Public Type DNSPacket
    Header As DNSHeader
    Question() As DNSQuestion
    Answer() As DNSRecord
    Authority() As DNSRecord
    Additional() As DNSRecord
End Type
Public Type SOA
    MName As String
    RName As String
    Serial As Double      '4 Bytes
    Refresh As Double     '4 Bytes
    Retry As Double       '4 Bytes
    Expire As Double      '4 Bytes
    Minimum As Double     '4 Bytes
End Type
Public Type MX
    Preference As Long    '2 Bytes
    Exchange As String
End Type
Public Type WKS
    Address As String
    Protocol As Byte      '1 Byte
    PortMap() As Boolean
End Type
Public Type HINFO
    CPU As String
    OS As String
End Type
Public Type MINFO
    RMailBX As String
    EMailBX As String
End Type
Public Type RP
    MBox_DName As String
    TXT_DName As String
End Type
Public Type AFSDB
    SubType As Long
    HostName As String
End Type
Public Type ISDN
    Address As String
    SA As String
End Type
Public Type RT
    Preference As Long
    Intermediate_Host As String
End Type
Public Type LOC
    Version As Byte
    Size As Byte
    Horiz_Pre As Byte
    Vert_Pre As Byte
    Latitude As Double
    Longitude As Double
    Altitude As Double
End Type

'Main Type Structure Filling Function

Public Function GetDNSInfo(Data) As DNSPacket
Dim NewRecord As DNSRecord
    With GetDNSInfo.Header
        .ID = (Data(0) * 256) + Data(1)

        .QR = Data(2) And 128 = 128
        .OpCode = ((Data(2) And 64) + (Data(2) And 32) + (Data(2) And 16) + (Data(2) And 8)) / 8
        .AA = (Data(2) And 4) = 4
        .TC = (Data(2) And 2) = 2
        .RD = (Data(2) And 1) = 1
        .RA = (Data(3) And 128) = 128
        .Z = ((Data(3) And 64) + (Data(3) And 32) + (Data(3) And 16)) / 16
        .RCode = (Data(3) And 8) + (Data(3) And 4) + (Data(3) And 2) + (Data(3) And 1)

        .QDCount = (Data(4) * 256) + Data(5)
        If .QDCount Then ReDim GetDNSInfo.Question(1 To .QDCount) As DNSQuestion

        .ANCount = (Data(6) * 256) + Data(7)
        If .ANCount Then ReDim GetDNSInfo.Answer(1 To .ANCount) As DNSRecord

        .NSCount = (Data(8) * 256) + Data(9)
        If .NSCount Then ReDim GetDNSInfo.Authority(1 To .NSCount) As DNSRecord

        .ARCount = (Data(10) * 256) + Data(11)
        If .ARCount Then ReDim GetDNSInfo.Additional(1 To .ARCount) As DNSRecord
    End With
    
    P = 12
    For X = 1 To GetDNSInfo.Header.QDCount
        With GetDNSInfo.Question(X)
            .QName = GetLabel(Data, P)

            .QType = (Data(P) * 256) + Data(P + 1)
            .QClass = (Data(P + 2) * 256) + Data(P + 3)
        End With
        P = P + 4
    Next

    For X = 1 To 3
        Select Case X
        Case 1: CurCount = GetDNSInfo.Header.ANCount
        Case 2: CurCount = GetDNSInfo.Header.NSCount
        Case 3: CurCount = GetDNSInfo.Header.ARCount
        End Select
        For Y = 1 To CurCount
            With NewRecord
                .RName = GetLabel(Data, P)

                .RType = (Data(P) * 256) + Data(P + 1)
                .RClass = (Data(P + 2) * 256) + Data(P + 3)
                .TTL = (Data(P + 4) * 16777216) + (Data(P + 5) * 65536) + (Data(P + 6) * 256) + Data(P + 7)
                .RDLength = (Data(P + 8) * 256) + Data(P + 9)

                P = P + 10
                Select Case .RType
                Case 1 'A
                    'Store the string version of the address
                    .RData = Data(P) & "." & Data(P + 1) & "." & Data(P + 2) & "." & Data(P + 3)
                    P = P + 4
                Case 2 To 4, 5, 7 To 9, 12 'NS, MD, MF, CNAME, MB, MG, MR, PTR
                    'Store the formatted Domain Name
                    .RData = GetLabel(Data, P)
                Case 16, 19 'TXT, X25
                    'Store the string
                    EndP = P + .RDLength
                    .RData = ""
                    While P < EndP
                        .RData = .RData & GetString(Data, P) & " "
                    Wend
                    .RData = Left(.RData, Len(.RData) - 1)
                Case 28 'IP6 Address
                    .RData = ""
                    For P = P To P + 15 Step 2
                        .RData = .RData & Right("000" & Hex(Data(P) * 256 + Data(P + 1)), 4) & ":"
                    Next
                    .RData = Left(.RData, Len(.RData) - 1)
                Case Else '6, 10, 11, 13, 14, 15, 16, 17, 20, 21, 29> 'SOA, NULL, WKS, HINFO, MINFO, MX, TXT, RP, ISDN, RT, LOC, Unknown
                    'Store a pointer to the data location
                    .RData = P
                    P = P + .RDLength
                End Select
            End With
            Select Case X
            Case 1: GetDNSInfo.Answer(Y) = NewRecord
            Case 2: GetDNSInfo.Authority(Y) = NewRecord
            Case 3: GetDNSInfo.Additional(Y) = NewRecord
            End Select
        Next
    Next
End Function

'Other Structure Fillers

Public Function GetSOA(Data, P) As SOA
Dim Dbl(4) As Double
    GetSOA.MName = GetLabel(Data, P)
    GetSOA.RName = Replace(GetLabel(Data, P), ".", "@", 1, 1)

    For X = 0 To 4
        Dbl(X) = (Data(P) * 16777216) + (Data(P + 1) * 65536) + ((Data(P + 2) * 256) + Data(P + 3))
        P = P + 4
    Next
    GetSOA.Serial = Dbl(0)
    GetSOA.Refresh = Dbl(1)
    GetSOA.Retry = Dbl(2)
    GetSOA.Expire = Dbl(3)
    GetSOA.Minimum = Dbl(4)
End Function
Public Function GetMX(Data, P) As MX
    GetMX.Preference = (Data(P) * 256) + Data(P + 1)
    P = P + 2
    GetMX.Exchange = GetLabel(Data, P)
End Function
Public Function GetWKS(Data, RDLength, P) As WKS
'Untested Function
    For P = P To P + 3
        GetWKS.Address = GetWKS.Address & Data(P) & "."
    Next
    GetWKS.Address = Left(GetWKS.Address, Len(GetWKS.Address) - 1)
    GetWKS.Protocol = Data(P)
    ReDim GetWKS.PortMap(((RDLength - 5) * 8) - 1) As Boolean
    For P = P + 1 To P + RDLength - 5
        For X = 7 To 0
            GetWKS.PortMap(Index) = Data(P) And (2 ^ X)
            Index = Index + 1
        Next
    Next
End Function
Public Function GetHINFO(Data, P) As HINFO
    GetHINFO.CPU = GetString(Data, P)
    GetHINFO.OS = GetString(Data, P)
End Function
Public Function GetMINFO(Data, P) As MINFO
    GetMINFO.RMailBX = GetLabel(Data, P)
    GetMINFO.EMailBX = GetLabel(Data, P)
End Function
Public Function GetRP(Data, P) As RP
    GetRP.MBox_DName = Replace(GetLabel(Data, P), ".", "@", 1, 1)
    GetRP.TXT_DName = GetLabel(Data, P)
End Function
Public Function GetAFSDB(Data, P) As AFSDB
    GetAFSDB.SubType = (Data(P) * 255) + Data(P + 1)
    GetAFSDB.HostName = GetLabel(Data, P + 2)
End Function
Public Function GetISDN(Data, P) As ISDN
    Length = (Data(P - 2) * 256) + Data(P - 1)
    If Length > Data(P) + 1 Then
        GetISDN.Address = GetString(Data, P)
        GetISDN.SA = GetString(Data, P)
    Else
        GetISDN.Address = GetString(Data, P)
    End If
End Function
Public Function GetRT(Data, P) As RT
    GetRT.Preference = (Data(P) * 255) + Data(P + 1)
    GetRT.Intermediate_Host = GetLabel(Data, P + 2)
End Function
Public Function GetLOC(Data, P) As LOC
Dim Dbl(2) As Double
    GetLOC.Version = Data(P)
    If GetLOC.Version = 0 Then
        GetLOC.Size = Data(P + 1)
        GetLOC.Horiz_Pre = Data(P + 2)
        GetLOC.Vert_Pre = Data(P + 3)

        For X = 0 To 2
            P = P + 4
            Dbl(X) = (Data(P) * 16777216) + (Data(P + 1) * 65536) + (Data(P + 2) * 256) + Data(P + 3)
        Next
        GetLOC.Latitude = Dbl(0)
        GetLOC.Longitude = Dbl(1)
        GetLOC.Altitude = Dbl(2)
    End If
End Function

'Private utility functions

Private Function GetLabel(Data, Pointer) As String
    For X = Pointer To UBound(Data)
        If Data(X) = 0 Then
            'Remove trailing period
            If GetLabel <> "" Then GetLabel = Left(GetLabel, Len(GetLabel) - 1)
            Pointer = IIf(TruePointer, TruePointer, X + 1)
            Exit Function
        'Label Length
        ElseIf Data(X) < 64 Then
            Label = ""
            For Y = X + 1 To X + Data(X)
                Label = Label & Chr(Data(Y))
            Next
            GetLabel = GetLabel & Label & "."
            X = X + Data(X)
        'Label Pointer
        ElseIf (Data(X) And 128) = 128 And (Data(X) And 64) = 64 Then
            If TruePointer = 0 Then TruePointer = X + 2
            X = (Data(X) - 192) * 256 + Data(X + 1) - 1
        End If
    Next
End Function
Private Function GetString(Data, P) As String
    For P = P + 1 To P + Data(P)
        GetString = GetString & Chr(Data(P))
    Next
End Function
