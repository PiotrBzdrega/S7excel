Attribute VB_Name = "Modul12"
'
' Part of Libnodave, a free communication libray for Siemens S7 200/300/400 via
' the MPI adapter 6ES7 972-0CA22-0XAC
' or  MPI adapter 6ES7 972-0CA23-0XAC
' or  TS adapter 6ES7 972-0CA33-0XAC
' or  MPI adapter 6ES7 972-0CA11-0XAC,
' IBH/MHJ-NetLink or CPs 243, 343 and 443
' or VIPA Speed7 with builtin ethernet support.
'
' (C) Thomas Hergenhahn (thomas.hergenhahn@web.de) 2005
'
' Libnodave is free software; you can redistribute it and/or modify
' it under the terms of the GNU Library General Public License as published by
' the Free Software Foundation; either version 2, or (at your option)
' any later version.
'
' Libnodave is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU Library General Public License
' along with Libnodave; see the file COPYING.  If not, write to
' the Free Software Foundation, 675 Mass Ave, Cambridge, MA 02139, USA.
'
'
'
'
'    Protocol types to be used with newInterface:
'
Private Const daveProtoMPI = 0      '  MPI for S7 300/400
Private Const daveProtoMPI2 = 1    '  MPI for S7 300/400, "Andrew's version"
Private Const daveProtoMPI3 = 2    '  MPI for S7 300/400, Step 7 Version, not yet implemented
Private Const daveProtoPPI = 10    '  PPI for S7 200
Private Const daveProtoAS511 = 20    '  S5 via programming interface
Private Const daveProtoS7online = 50    '  S7 using Siemens libraries & drivers for transport
Private Const daveProtoISOTCP = 122 '  ISO over TCP
Private Const daveProtoISOTCP243 = 123 '  ISO o?ver TCP with CP243
Private Const daveProtoMPI_IBH = 223   '  MPI with IBH NetLink MPI to ethernet gateway */
Private Const daveProtoPPI_IBH = 224   '  PPI with IBH NetLink PPI to ethernet gateway */
Private Const daveProtoUserTransport = 255 '  Libnodave will pass the PDUs of S7 Communication to user defined call back functions.
'
'    ProfiBus speed constants:
'
Private Const daveSpeed9k = 0
Private Const daveSpeed19k = 1
Private Const daveSpeed187k = 2
Private Const daveSpeed500k = 3
Private Const daveSpeed1500k = 4
Private Const daveSpeed45k = 5
Private Const daveSpeed93k = 6
'
'    S7 specific constants:
'
Private Const daveBlockType_OB = "8"
Private Const daveBlockType_DB = "A"
Private Const daveBlockType_SDB = "B"
Private Const daveBlockType_FC = "C"
Private Const daveBlockType_SFC = "D"
Private Const daveBlockType_FB = "E"
Private Const daveBlockType_SFB = "F"
'
' Use these constants for parameter "area" in daveReadBytes and daveWriteBytes
'
Private Const daveSysInfo = &H3      '  System info of 200 family
Private Const daveSysFlags = &H5   '  System flags of 200 family
Private Const daveAnaIn = &H6      '  analog inputs of 200 family
Private Const daveAnaOut = &H7     '  analog outputs of 200 family
Private Const daveP = &H80          ' direct access to peripheral adresses
Private Const daveInputs = &H81
Private Const daveOutputs = &H82
Private Const daveFlags = &H83
Private Const daveDB = &H84 '  data blocks
Private Const daveDI = &H85  '  instance data blocks
Private Const daveV = &H87      ' don't know what it is
Private Const daveCounter = 28  ' S7 counters
Private Const daveTimer = 29    ' S7 timers
Private Const daveCounter200 = 30       ' IEC counters (200 family)
Private Const daveTimer200 = 31         ' IEC timers (200 family)
'
Private Const daveOrderCodeSize = 21    ' Length of order code (MLFB number)
'
'    Library specific:
'
'
'    Result codes. Genarally, 0 means ok,
'    >0 are results (also errors) reported by the PLC
'    <0 means error reported by library code.
'
Private Const daveResOK = 0                        ' means all ok
Private Const daveResNoPeripheralAtAddress = 1     ' CPU tells there is no peripheral at address
Private Const daveResMultipleBitsNotSupported = 6  ' CPU tells it does not support to read a bit block with a
                                                   ' length other than 1 bit.
Private Const daveResItemNotAvailable200 = 3       ' means a a piece of data is not available in the CPU, e.g.
                                                   ' when trying to read a non existing DB or bit bloc of length<>1
                                                   ' This code seems to be specific to 200 family.
Private Const daveResItemNotAvailable = 10         ' means a a piece of data is not available in the CPU, e.g.
                                                   ' when trying to read a non existing DB
Private Const daveAddressOutOfRange = 5            ' means the data address is beyond the CPUs address range
Private Const daveWriteDataSizeMismatch = 7        ' means the write data size doesn't fit item size
Private Const daveResCannotEvaluatePDU = -123
Private Const daveResCPUNoData = -124
Private Const daveUnknownError = -125
Private Const daveEmptyResultError = -126
Private Const daveEmptyResultSetError = -127
Private Const daveResUnexpectedFunc = -128
Private Const daveResUnknownDataUnitSize = -129
Private Const daveResShortPacket = -1024
Private Const daveResTimeout = -1025
'
'    Max number of bytes in a single message.
'
Private Const daveMaxRawLen = 2048
'
'    Some definitions for debugging:
'
Private Const daveDebugRawRead = &H1            ' Show the single bytes received
Private Const daveDebugSpecialChars = &H2       ' Show when special chars are read
Private Const daveDebugRawWrite = &H4           ' Show the single bytes written
Private Const daveDebugListReachables = &H8     ' Show the steps when determine devices in MPI net
Private Const daveDebugInitAdapter = &H10       ' Show the steps when Initilizing the MPI adapter
Private Const daveDebugConnect = &H20           ' Show the steps when connecting a PLC
Private Const daveDebugPacket = &H40
Private Const daveDebugByte = &H80
Private Const daveDebugCompare = &H100
Private Const daveDebugExchange = &H200
Private Const daveDebugPDU = &H400      ' debug PDU handling
Private Const daveDebugUpload = &H800   ' debug PDU loading program blocks from PLC
Private Const daveDebugMPI = &H1000
Private Const daveDebugPrintErrors = &H2000     ' Print error messages
Private Const daveDebugPassive = &H4000
Private Const daveDebugErrorReporting = &H8000
Private Const daveDebugOpen = &H8000
Private Const daveDebugAll = &H1FFFF
'
'    Set and read debug level:
'


Private Declare PtrSafe Sub FortranCall Lib "libnodave.dll" ()
Private Declare PtrSafe Sub daveSetDebug Lib "libnodave.dll" (ByVal level As Long)
Private Declare PtrSafe Function daveGetDebug Lib "libnodave.dll" () As Long
'
' You may wonder what sense it might make to set debug level, as you cannot see
' messages when you opened excel or some VB application from Windows GUI.
' You can invoke Excel from the console or from a batch file with:
' <myPathToExcel>\Excel.Exe <MyPathToXLS-File>VBATest.XLS >ExcelOut
' This will start Excel with VBATest.XLS and all debug messages (and a few from Excel itself)
' go into the file ExcelOut.
'
'    Error code to message string conversion:
'    Call this function to get an explanation for error codes returned by other functions.
'
'
' The folowing doesn't work properly. A VB string is something different from a pointer to char:
'
' Private Declare Function daveStrerror Lib "libnodave.dll" Alias "daveStrerror" (ByVal en As Long) As String
'
Private Declare PtrSafe Function daveInternalStrerror Lib "libnodave.dll" Alias "daveStrerror" (ByVal en As Long) As Long
' So, I added another function to libnodave wich copies the text into a VB String.
' This function is still not useful without some code araound it, so I call it "internal"
Private Declare PtrSafe Sub daveStringCopy Lib "libnodave.dll" (ByVal internalPointer As Long, ByVal s As String)
'
' Setup a new interface structure using a handle to an open port or socket:
'
Private Declare PtrSafe Function daveNewInterface Lib "libnodave.dll" (ByVal fd1 As Long, ByVal fd2 As Long, ByVal name As String, ByVal localMPI As Long, ByVal protocol As Long, ByVal speed As Long) As Long
'
' Setup a new connection structure using an initialized daveInterface and PLC's MPI address.
' Note: The parameter di must have been obtained from daveNewinterface.
'
Private Declare PtrSafe Function daveNewConnection Lib "libnodave.dll" (ByVal di As Long, ByVal mpi As Long, ByVal Rack As Long, ByVal Slot As Long) As Long
'
'    PDU handling:
'    PDU is the central structure present in S7 communication.
'    It is composed of a 10 or 12 byte header,a parameter block and a data block.
'    When reading or writing values, the data field is itself composed of a data
'    header followed by payload data
'
'    retrieve the answer:
'    Note: The parameter dc must have been obtained from daveNewConnection.
'
Private Declare PtrSafe Function daveGetResponse Lib "libnodave.dll" (ByVal dc As Long) As Long
'
'    send PDU to PLC
'    Note: The parameter dc must have been obtained from daveNewConnection,
'          The parameter pdu must have been obtained from daveNewPDU.
'
Private Declare PtrSafe Function daveSendMessage Lib "libnodave.dll" (ByVal dc As Long, ByVal pdu As Long) As Long
'******
'
'Utilities:
'
'****
'*
'    Hex dump PDU:
'
Private Declare PtrSafe Sub daveDumpPDU Lib "libnodave.dll" (ByVal pdu As Long)
'
'    Hex dump. Write the name followed by len bytes written in hex and a newline:
'
Private Declare PtrSafe Sub daveDump Lib "libnodave.dll" (ByVal name As String, ByVal pdu As Long, ByVal length As Long)
'
'    names for PLC objects. This is again the intenal function. Use the wrapper code below.
'
Private Declare PtrSafe Function daveInternalAreaName Lib "libnodave.dll" Alias "daveAreaName" (ByVal en As Long) As Long
Private Declare PtrSafe Function daveInternalBlockName Lib "libnodave.dll" Alias "daveBlockName" (ByVal en As Long) As Long
'
'   swap functions. They change the byte order, if byte order on the computer differs from
'   PLC byte order:
'
Private Declare PtrSafe Function daveSwapIed_16 Lib "libnodave.dll" (ByVal x As Long) As Long
Private Declare PtrSafe Function daveSwapIed_32 Lib "libnodave.dll" (ByVal x As Long) As Long
'
'    Data conversion convenience functions. The older set has been removed.
'    Newer conversion routines. As the terms WORD, INT, INTEGER etc have different meanings
'    for users of different programming languages and compilers, I choose to provide a new
'    set of conversion routines named according to the bit length of the value used. The 'U'
'    or 'S' stands for unsigned or signed.
'
'
'    Get a value from the position b points to. B is typically a pointer to a buffer that has
'    been filled with daveReadBytes:
'
Private Declare PtrSafe Function toPLCfloat Lib "libnodave.dll" (ByVal f As Single) As Single
Private Declare PtrSafe Function daveToPLCfloat Lib "libnodave.dll" (ByVal f As Single) As Long
'
' Copy and convert value of 8,16,or 32 bit, signed or unsigned at position pos
' from internal buffer:
'
Private Declare PtrSafe Function daveGetS8from Lib "libnodave.dll" (ByRef buffer As Byte) As Long
Private Declare PtrSafe Function daveGetU8from Lib "libnodave.dll" (ByRef buffer As Byte) As Long
Private Declare PtrSafe Function daveGetS16from Lib "libnodave.dll" (ByRef buffer As Byte) As Long
Private Declare PtrSafe Function daveGetU16from Lib "libnodave.dll" (ByRef buffer As Byte) As Long
Private Declare PtrSafe Function daveGetS32from Lib "libnodave.dll" (ByRef buffer As Byte) As Long
'
' Is there an unsigned long? Or a longer integer than long? This doesn't work.
' Private Declare Function daveGetU32from Lib "libnodave.dll" (ByRef buffer As Byte) As Long
'
Private Declare PtrSafe Function daveGetFloatfrom Lib "libnodave.dll" (ByRef buffer As Byte) As Single
'
' Copy and convert a value of 8,16,or 32 bit, signed or unsigned from internal buffer. These
' functions increment an internal buffer position. This buffer position is set to zero by
' daveReadBytes, daveReadBits, daveReadSZL.
'
Private Declare PtrSafe Function daveGetS8 Lib "libnodave.dll" (ByVal dc As Long) As Long
Private Declare PtrSafe Function daveGetU8 Lib "libnodave.dll" (ByVal dc As Long) As Long
Private Declare PtrSafe Function daveGetS16 Lib "libnodave.dll" (ByVal dc As Long) As Long
Private Declare PtrSafe Function daveGetU16 Lib "libnodave.dll" (ByVal dc As Long) As Long
Private Declare PtrSafe Function daveGetS32 Lib "libnodave.dll" (ByVal dc As Long) As Long
'
' Is there an unsigned long? Or a longer integer than long? This doesn't work.
'Private Declare Function daveGetU32 Lib "libnodave.dll" (ByVal dc As Long) As Long
Private Declare PtrSafe Function daveGetFloat Lib "libnodave.dll" (ByVal dc As Long) As Single
'
' Read a value of 8,16,or 32 bit, signed or unsigned at position pos from internal buffer:
'
Private Declare PtrSafe Function daveGetS8At Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
Private Declare PtrSafe Function daveGetU8At Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
Private Declare PtrSafe Function daveGetS16At Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
Private Declare PtrSafe Function daveGetU16At Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
Private Declare PtrSafe Function daveGetS32At Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
'
' Is there an unsigned long? Or a longer integer than long? This doesn't work.
'Private Declare Function daveGetU32At Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
Private Declare PtrSafe Function daveGetFloatAt Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Single
'
' Copy and convert a value of 8,16,or 32 bit, signed or unsigned into a buffer. The buffer
' is usually used by daveWriteBytes, daveWriteBits later.
'
Private Declare PtrSafe Function davePut8 Lib "libnodave.dll" (ByRef buffer As Byte, ByVal value As Long) As Long
Private Declare PtrSafe Function davePut16 Lib "libnodave.dll" (ByRef buffer As Byte, ByVal value As Long) As Long
Private Declare PtrSafe Function davePut32 Lib "libnodave.dll" (ByRef buffer As Byte, ByVal value As Long) As Long
Private Declare PtrSafe Function davePutFloat Lib "libnodave.dll" (ByRef buffer As Byte, ByVal value As Single) As Long
'
' Copy and convert a value of 8,16,or 32 bit, signed or unsigned to position pos of a buffer.
' The buffer is usually used by daveWriteBytes, daveWriteBits later.
'
Private Declare PtrSafe Function davePut8At Lib "libnodave.dll" (ByRef buffer As Byte, ByVal pos As Long, ByVal value As Long) As Long
Private Declare PtrSafe Function davePut16At Lib "libnodave.dll" (ByRef buffer As Byte, ByVal pos As Long, ByVal value As Long) As Long
Private Declare PtrSafe Function davePut32At Lib "libnodave.dll" (ByRef buffer As Byte, ByVal pos As Long, ByVal value As Long) As Long
Private Declare PtrSafe Function davePutFloatAt Lib "libnodave.dll" (ByRef buffer As Byte, ByVal pos As Long, ByVal value As Single) As Long
'
' Takes a timer value and converts it into seconds:
'
Private Declare PtrSafe Function daveGetSeconds Lib "libnodave.dll" (ByVal dc As Long) As Single
Private Declare PtrSafe Function daveGetSecondsAt Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Single
'
' Takes a counter value and converts it to integer:
'
Private Declare PtrSafe Function daveGetCounterValue Lib "libnodave.dll" (ByVal dc As Long) As Long
Private Declare PtrSafe Function daveGetCounterValueAt Lib "libnodave.dll" (ByVal dc As Long, ByVal pos As Long) As Long
'
' Get the order code (MLFB number) from a PLC. Does NOT work with 200 family.
'
Private Declare PtrSafe Function daveGetOrderCode Lib "libnodave.dll" (ByVal en As Long, ByRef buffer As Byte) As Long
'
' Connect to a PLC.
'
Private Declare PtrSafe Function daveConnectPLC Lib "libnodave.dll" (ByVal dc As Long) As Long
'
'
' Read a value or a block of values from PLC.
'
Private Declare PtrSafe Function daveReadBytes Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal areaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByVal buffer As Long) As Long
'
' Read a long block of values from PLC. Long means too long to transport in a single PDU.
'
Private Declare PtrSafe Function daveManyReadBytes Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal areaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByVal buffer As Long) As Long
'
' Write a value or a block of values to PLC.
'
Private Declare PtrSafe Function daveWriteBytes Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal areaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByRef buffer As Byte) As Long
'
' Write a long block of values to PLC. Long means too long to transport in a single PDU.
'
Private Declare PtrSafe Function daveWriteManyBytes Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal areaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByRef buffer As Byte) As Long
'
' Read a bit from PLC. numBytes must be exactly one with all PLCs tested.
' Start is calculated as 8*byte number+bit number.
'
Private Declare PtrSafe Function daveReadBits Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal areaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByVal buffer As Long) As Long
'
' Write a bit to PLC. numBytes must be exactly one with all PLCs tested.
'
Private Declare PtrSafe Function daveWriteBits Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal areaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByRef buffer As Long) As Long
'
' Set a bit in PLC to 1. pb: deleted start parameter
'
Private Declare PtrSafe Function daveSetBit Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal areaNumber As Long, ByVal byteAddress As Long, ByVal bitAddress As Long) As Long
'
' Set a bit in PLC to 0. pb: deleted start parameter
'
Private Declare PtrSafe Function daveClrBit Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal areaNumber As Long, ByVal byteAddress As Long, ByVal bitAddress As Long) As Long
'
' Read a diagnostic list (SZL) from PLC. Does NOT work with 200 family.
'
Private Declare PtrSafe Function daveReadSZL Lib "libnodave.dll" (ByVal dc As Long, ByVal ID As Long, ByVal index As Long, ByRef buffer As Byte, ByVal buflen As Long) As Long
'
Private Declare PtrSafe Function daveListBlocksOfType Lib "libnodave.dll" (ByVal dc As Long, ByVal typ As Long, ByRef buffer As Byte) As Long
Private Declare PtrSafe Function daveListBlocks Lib "libnodave.dll" (ByVal dc As Long, ByRef buffer As Byte) As Long
Private Declare PtrSafe Function internalDaveGetBlockInfo Lib "libnodave.dll" Alias "daveGetBlockInfo" (ByVal dc As Long, ByRef buffer As Byte, ByVal btype As Long, ByVal number As Long) As Long
'
Private Declare PtrSafe Function daveGetProgramBlock Lib "libnodave.dll" (ByVal dc As Long, ByVal blockType As Long, ByVal number As Long, ByRef buffer As Byte, ByRef length As Long) As Long
'
' Start or Stop a PLC:
'
Private Declare PtrSafe Function daveStart Lib "libnodave.dll" (ByVal dc As Long) As Long
Private Declare PtrSafe Function daveStop Lib "libnodave.dll" (ByVal dc As Long) As Long
'
' Set outputs (digital or analog ones) of an S7-200 that is in stop mode:
'
Private Declare PtrSafe Function daveForce200 Lib "libnodave.dll" (ByVal dc As Long, ByVal area As Long, ByVal start As Long, ByVal value As Long) As Long
'
' Initialize a multivariable read request.
' The parameter PDU must have been obtained from daveNew PDU:
'
Private Declare PtrSafe Sub davePrepareReadRequest Lib "libnodave.dll" (ByVal dc As Long, ByVal pdu As Long)
'
' Add a new variable to a prepared request:
'
Private Declare PtrSafe Sub daveAddVarToReadRequest Lib "libnodave.dll" (ByVal pdu As Long, ByVal area As Long, ByVal areaNumber As Long, ByVal start As Long, ByVal numBytes As Long)
'
' Executes the entire request:
'
Private Declare PtrSafe Function daveExecReadRequest Lib "libnodave.dll" (ByVal dc As Long, ByVal pdu As Long, ByVal rs As Long) As Long
'
' Use the n-th result. This lets the functions daveGet<data type> work on that part of the
' internal buffer that contains the n-th result:
'
Private Declare PtrSafe Function daveUseResult Lib "libnodave.dll" (ByVal dc As Long, ByVal rs As Long, ByVal resultNumber As Long) As Long
'
' Frees the memory occupied by single results in the result structure. After that, you can reuse
' the resultSet in another call to daveExecReadRequest.
'
Private Declare PtrSafe Sub daveFreeResults Lib "libnodave.dll" (ByVal rs As Long)
'
' Adds a new bit variable to a prepared request. As with daveReadBits, numBytes must be one for
' all tested PLCs.
'
Private Declare PtrSafe Sub daveAddBitVarToReadRequest Lib "libnodave.dll" (ByVal pdu As Long, ByVal area As Long, ByVal areaNumber As Long, ByVal start As Long, ByVal numBytes As Long)
'
' Initialize a multivariable write request.
' The parameter PDU must have been obtained from daveNew PDU:
'
Private Declare PtrSafe Sub davePrepareWriteRequest Lib "libnodave.dll" (ByVal dc As Long, ByVal pdu As Long)
'
' Add a new variable to a prepared write request:
'
Private Declare PtrSafe Sub daveAddVarToWriteRequest Lib "libnodave.dll" (ByVal pdu As Long, ByVal area As Long, ByVal areaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByRef buffer As Byte)
'
' Add a new bit variable to a prepared write request:
'
Private Declare PtrSafe Sub daveAddBitVarToWriteRequest Lib "libnodave.dll" (ByVal pdu As Long, ByVal area As Long, ByVal areaNumber As Long, ByVal start As Long, ByVal numBytes As Long, ByRef buffer As Byte)
'
' Execute the entire write request:
'
Private Declare PtrSafe Function daveExecWriteRequest Lib "libnodave.dll" (ByVal dc As Long, ByVal pdu As Long, ByVal rs As Long) As Long
'
' Initialize an MPI Adapter or NetLink Ethernet MPI gateway.
' While some protocols do not need this, I recommend to allways use it. It will do nothing if
' the protocol doesn't need it. But you can change protocols without changing your program code.
'
Private Declare PtrSafe Function daveInitAdapter Lib "libnodave.dll" (ByVal di As Long) As Long
'
' Disconnect from a PLC. While some protocols do not need this, I recommend to allways use it.
' It will do nothing if the protocol doesn't need it. But you can change protocols without
' changing your program code.
'
Private Declare PtrSafe Function daveDisconnectPLC Lib "libnodave.dll" (ByVal dc As Long) As Long
'
'
' Disconnect from an MPI Adapter or NetLink Ethernet MPI gateway.
' While some protocols do not need this, I recommend to allways use it.
' It will do nothing if the protocol doesn't need it. But you can change protocols without
' changing your program code.
'
Private Declare PtrSafe Function daveDisconnectAdapter Lib "libnodave.dll" (ByVal dc As Long) As Long
'
'
' List nodes on an MPI or Profibus Network:
'
Private Declare PtrSafe Function daveListReachablePartners Lib "libnodave.dll" (ByVal dc As Long, ByRef buffer As Byte) As Long
'
'
' Set/change the timeout for an interface:
'
Private Declare PtrSafe Sub daveSetTimeout Lib "libnodave.dll" (ByVal di As Long, ByVal maxTime As Long)
'
' Read the timeout setting for an interface:
'
Private Declare PtrSafe Function daveGetTimeout Lib "libnodave.dll" (ByVal di As Long)
'
' Get the name of an interface. Do NOT use this, but the wrapper function defined below!
'
Private Declare PtrSafe Function daveInternalGetName Lib "libnodave.dll" Alias "daveGetName" (ByVal en As Long) As Long
'
' Get the MPI address of a connection.
'
Private Declare PtrSafe Function daveGetMPIAdr Lib "libnodave.dll" (ByVal dc As Long) As Long
'
' Get the length (in bytes) of the last data received on a connection.
'
Private Declare PtrSafe Function daveGetAnswLen Lib "libnodave.dll" (ByVal dc As Long) As Long
'
' Get the maximum length of a communication packet (PDU).
' This value depends on your CPU and connection type. It is negociated in daveConnectPLC.
' A simple read can read MaxPDULen-18 bytes.
'
Private Declare PtrSafe Function daveGetMaxPDULen Lib "libnodave.dll" (ByVal dc As Long) As Long
'
' Reserve memory for a resultSet and get a handle to it:
'
Private Declare PtrSafe Function daveNewResultSet Lib "libnodave.dll" () As Long
'
' Destroy handles to daveInterface, daveConnections, PDUs and resultSets
' Free the memory reserved for them.
'
Private Declare PtrSafe Sub daveFree Lib "libnodave.dll" (ByVal item As Long)
'
' Reserve memory for a PDU and get a handle to it:
'
Private Declare PtrSafe Function daveNewPDU Lib "libnodave.dll" () As Long
'
' Get the error code of the n-th single result in a result set:
'
Private Declare PtrSafe Function daveGetErrorOfResult Lib "libnodave.dll" (ByVal resultSet As Long, ByVal resultNumber As Long) As Long
'
Private Declare PtrSafe Function daveForceDisconnectIBH Lib "libnodave.dll" (ByVal di As Long, ByVal src As Long, ByVal dest As Long, ByVal mpi As Long) As Long
'
' Helper functions to open serial ports and IP connections. You can use others if you want and
' pass their results to daveNewInterface.
'
' Open a serial port using name, baud rate and parity. Everything else is set automatically:
'
Private Declare PtrSafe Function setPort Lib "libnodave.dll" (ByVal portName As String, ByVal baudrate As String, ByVal parity As Byte) As Long
'
' Open a TCP/IP connection using port number (1099 for NetLink, 102 for ISO over TCP) and
' IP address. You must use an IP address, NOT a hostname!
'
Private Declare PtrSafe Function openSocket Lib "libnodave.dll" (ByVal port As Long, ByVal peer As String) As Long
'
' Open an access oint. This is a name in you can add in the "set Programmer/PLC interface" dialog.
' To the access point, you can assign an interface like MPI adapter, CP511 etc.
'
Private Declare PtrSafe Function openS7online Lib "libnodave.dll" (ByVal peer As String) As Long
'
' Close connections and serial ports opened with above functions:
'
Private Declare PtrSafe Function closePort Lib "libnodave.dll" (ByVal fh As Long) As Long
'
' Close sockets opened with above functions:
'
Private Declare PtrSafe Function closeSocket Lib "libnodave.dll" (ByVal fh As Long) As Long
'
' Close handle opened by opens7online:
'
Private Declare PtrSafe Function closeS7online Lib "libnodave.dll" (ByVal fh As Long) As Long
'
' Read Clock time from PLC:
'
Private Declare PtrSafe Function daveReadPLCTime Lib "libnodave.dll" (ByVal dc As Long) As Long
'
' set clock to a value given by user
'
Private Declare PtrSafe Function daveSetPLCTime Lib "libnodave.dll" (ByVal dc As Long, ByRef timestamp As Byte) As Long
'
' set clock to PC system clock:
'
Private Declare PtrSafe Function daveSetPLCTimeToSystime Lib "libnodave.dll" (ByVal dc As Long) As Long
'
'       BCD conversions:
'
Private Declare PtrSafe Function daveToBCD Lib "libnodave.dll" (ByVal dc As Long) As Long
Private Declare PtrSafe Function daveFromBCD Lib "libnodave.dll" (ByVal dc As Long) As Long
'
' Here comes the wrapper code for functions returning strings:
'
Private Function daveStrError(ByVal code As Long) As String
    x$ = String$(256, 0)            'create a string of sufficient capacity
    ip = daveInternalStrerror(code)    ' have the text for code copied in
    Call daveStringCopy(ip, x$)    ' have the text for code copied in
    x$ = Left$(x$, InStr(x$, Chr$(0)) - 1) ' adjust the length
    daveStrError = x$                       ' and return result
End Function

Private Function daveAreaName(ByVal code As Long) As String
    x$ = String$(256, 0)            'create a string of sufficient capacity
    ip = daveInternalAreaName(code)    ' have the text for code copied in
    Call daveStringCopy(ip, x$)    ' have the text for code copied in
    x$ = Left$(x$, InStr(x$, Chr$(0)) - 1) ' adjust the length
    daveAreaName = x$                       ' and return result
End Function
Private Function daveBlockName(ByVal code As Long) As String
    x$ = String$(256, 0)            'create a string of sufficient capacity
    ip = daveInternalBlockName(code)    ' have the text for code copied in
    Call daveStringCopy(ip, x$)    ' have the text for code copied in
    x$ = Left$(x$, InStr(x$, Chr$(0)) - 1) ' adjust the length
    daveBlockName = x$                       ' and return result
End Function
Private Function daveGetName(ByVal di As Long) As String
    x$ = String$(256, 0)            'create a string of sufficient capacity
    ip = daveInternalGetName(di)    ' have the text for code copied in
    Call daveStringCopy(ip, x$)    ' have the text for code copied in
    x$ = Left$(x$, InStr(x$, Chr$(0)) - 1) ' adjust the length
    daveGetName = x$                       ' and return result
End Function
Private Function daveGetBlockInfo(ByVal di As Long) As Byte
    x$ = String$(256, 0)            'create a string of sufficient capacity
    ip = daveInternalGetName(di)    ' have the text for code copied in
    Call daveStringCopy(ip, x$)    ' have the text for code copied in
    x$ = Left$(x$, InStr(x$, Chr$(0)) - 1) ' adjust the length
    daveGetName = x$                       ' and return result
End Function



'
' This initialization is used in all test programs. In a real program, where you would
' want to read again and again, keep the dc and di until your program terminates.
'
Private Function Initialize(ByRef ph As Long, ByRef di As Long, ByRef dc As Long)

ph = 0
di = 0
dc = 0
Initialize = -1
res = -1
peer$ = ActiveWorkbook.Worksheets("VarTab").Cells(2, 9)
ph = openSocket(102, peer$)    ' for ISO over TCP
If (ph > 0) Then
di = daveNewInterface(ph, ph, "IF1", 0, daveProtoISOTCP, daveSpeed500k)
'    Call daveSetTimeout(di, 500000)
    res = daveInitAdapter(di)
    If res = 0 Then
        MpiPpi = Cells(6, 5)
'
' with ISO over TCP, set correct values for rack and slot of the CPU
'
        dc = daveNewConnection(di, MpiPpi, ActiveWorkbook.Worksheets("VarTab").Cells(3, 9), ActiveWorkbook.Worksheets("VarTab").Cells(4, 9))
        res = daveConnectPLC(dc)
        If res = 0 Then
            Initialize = 0
        End If
    End If
End If
End Function
'
' Disconnect from PLC, disconnect from Adapter, close the serial interface or TCP/IP socket
'
Private Sub cleanUp(ByRef ph As Long, ByRef di As Long, ByRef dc As Long)
If dc <> 0 Then
    res = daveDisconnectPLC(dc)
    Call daveFree(dc)
    dc = 0
End If
If di <> 0 Then
    res = daveDisconnectAdapter(di)
    Call daveFree(di)
    di = 0
End If
If ph <> 0 Then
    res = closePort(ph)
'   res = closeSocket(ph)
    ph = 0
End If
End Sub
'--------------------------------Read data on PLC--------------------------------------
Sub readFromPLC()
    Dim ph As Long, di As Long, dc As Long, iRow As Integer, dbnum As String, addrOffset As String, addrBit As String
    Dim TagName As String, TagAdress As String, TagType_array() As String
    'pb area variables for I/O
    Dim area As Long, areaNum As Long
    areaNum = 0
    area = 0 'to check if we want to get i/o or db variable
    'pb
    iRow = 3
    res = Initialize(ph, di, dc)
    If res = 0 Then
        Do Until IsEmpty(ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 3))
            TagAdress = ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 3).value
            TagType_array = Split(ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 3).value, ".")
            
            'pb assign area for determine variable
            If InStr(TagType_array(0), "DB") > 0 Then
                'datablock variable
                dbnum = Replace(TagType_array(0), "DB", "")
            ElseIf InStr(TagType_array(0), "I") > 0 Then
                'inupt'
                area = daveInputs
            ElseIf InStr(TagType_array(0), "Q") > 0 Then
                'outpt'
                area = daveOutputs
            End If
            'pb
            

'--------------------------------Read bit data--------------------------------------
            If UBound(TagType_array) > 1 Then
               
                    addrOffset = Replace(TagType_array(1), "DBX", "")
                    addrBit = TagType_array(2)
                    res2 = daveReadBytes(dc, daveDB, dbnum, addrOffset, 1, 0)
                    If res2 = 0 Then
                        Dim bfbyte As Byte
                        Dim bitStat As Integer
                        Dim bitPos As Byte
                        bitPos = CByte(addrBit)
                        bfbyte = daveGetU8(dc)                                               'Copy and convert a value of 8 bit, signed or unsigned from internal buffer
                        bitStat = bfbyte And 2 ^ bitPos
    'Convert byte to bit
                        If bitStat > 0 Then
                            ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = True
                        Else
                            ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = False
                        End If
                    Else
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = "#####"
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4).Interior.Color = RGB(255, 0, 0)
                    End If
            Else
'--------------------------------Read bit i/o data--------------------------------------

                'pb we searching for input
                If area = daveInputs Then
                    addrOffset = Replace(TagType_array(0), "I", "")
                    addrBit = TagType_array(1)
                    res2 = daveReadBytes(dc, area, areaNum, addrOffset, 1, 0)
                    If res2 = 0 Then
                        Dim bfbyteI As Byte
                        Dim bitStatI As Integer
                        Dim bitPosI As Byte
                        bitPosI = CByte(addrBit)
                        bfbyteI = daveGetU8(dc)                                               'Copy and convert a value of 8 bit, signed or unsigned from internal buffer
                        bitStatI = bfbyteI And 2 ^ bitPosI
    'Convert byte to bit
                        If bitStatI > 0 Then
                            ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = True
                        Else
                            ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = False
                        End If
                    Else
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = "#####"
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4).Interior.Color = RGB(255, 0, 0)
                    End If
                
                'pb we searching for output
                ElseIf area = daveOutputs Then
                    addrOffset = Replace(TagType_array(0), "Q", "")
                    addrBit = TagType_array(1)
                    res2 = daveReadBytes(dc, area, areaNum, addrOffset, 1, 0)
                    If res2 = 0 Then
                        Dim bfbyteQ As Byte
                        Dim bitStatQ As Integer
                        Dim bitPosQ As Byte
                        bitPosQ = CByte(addrBit)
                        bfbyteQ = daveGetU8(dc)                                               'Copy and convert a value of 8 bit, signed or unsigned from internal buffer
                        bitStatQ = bfbyteQ And 2 ^ bitPosQ
    'Convert byte to bit
                        If bitStatQ > 0 Then
                            ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = True
                        Else
                            ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = False
                        End If
                    Else
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = "#####"
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4).Interior.Color = RGB(255, 0, 0)
                    End If
'--------------------------------read dword data--------------------------------------
                ElseIf InStr(TagType_array(1), "DBDW") > 0 Then
                    addrOffset = Replace(TagType_array(1), "DBDW", "")
                    res2 = daveReadBytes(dc, daveDB, dbnum, addrOffset, 4, 0)
                    If res2 = 0 Then
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = daveGetS32(dc) 'Copy and convert a value of 32 bit, signed or unsigned from internal buffer
                    Else
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = "#####"
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4).Interior.Color = RGB(255, 0, 0)
                    End If
'--------------------------------Read real data--------------------------------------
                ElseIf InStr(TagType_array(1), "DBD") > 0 Then
                    addrOffset = Replace(TagType_array(1), "DBD", "")
                    res2 = daveReadBytes(dc, daveDB, dbnum, addrOffset, 4, 0)
                    If res2 = 0 Then
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = daveGetFloat(dc) 'Copy and convert a value of 32 bit, signed or unsigned from internal buffer
                    Else
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = "#####"
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4).Interior.Color = RGB(255, 0, 0)
                    End If

'--------------------------------read word data--------------------------------------
                ElseIf InStr(TagType_array(1), "DBW") > 0 Then
                    addrOffset = Replace(TagType_array(1), "DBW", "")
                    res2 = daveReadBytes(dc, daveDB, dbnum, addrOffset, 2, 0)
                    If res2 = 0 Then
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = daveGetU16(dc) 'Copy and convert a value of 16 bit, signed or unsigned from internal buffer
                    Else
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = "#####"
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4).Interior.Color = RGB(255, 0, 0)
                    End If
'--------------------------------read byte data--------------------------------------
                ElseIf InStr(TagType_array(1), "DBB") > 0 Then
                    addrOffset = Replace(TagType_array(1), "DBB", "")
                    res2 = daveReadBytes(dc, daveDB, dbnum, addrOffset, 1, 0)
                    If res2 = 0 Then
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = daveGetU8(dc) 'Copy and convert a value of 8 bit, signed or unsigned from internal buffer
                    Else
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4) = "#####"
                        ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 4).Interior.Color = RGB(255, 0, 0)
                    End If
                End If
            End If
            iRow = iRow + 1
        Loop
    Else
        MsgBox "No route PLC, check connection and settings"
    End If
        Call cleanUp(ph, di, dc)
End Sub
'--------------------------------Write data on PLC--------------------------------------
Sub WriteFromPLC()
    Dim ph As Long, di As Long, dc As Long, iRow As Integer, dbnum As String, addrOffset As String, addrBit As String
    Dim TagName As String, TagAdress As String, TagType_array() As String, TagCompare As String
    Dim buffer As Byte, buffer1 As Byte, buffer2 As Byte, buffer3 As Byte
    Dim bfbyte As Byte, bitStat As Integer, bitPos As Byte
    Dim areaNum As Long
    areaNum = 0 'variable for real i/o
    
    iRow = 3
    addrBit2 = 0
    
    res = Initialize(ph, di, dc)
    If res = 0 Then
        Do Until IsEmpty(ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 3))
            TagAdress = ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 3).value
            TagType_array = Split(ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 3).value, ".")
            TagType_array2 = Split(ActiveWorkbook.Worksheets("VarTab").Cells(iRow + 1, 3).value, ".")
            
            dbnum = Replace(TagType_array(0), "DB", "")
'--------------------------------Write bit data--------------------------------------
            If UBound(TagType_array) > 1 Then
                addrOffset = Replace(TagType_array(1), "DBX", "")
                addrBit = TagType_array(2)
'Check if are in the same DB to complet the byte
                If IsEmpty(ActiveWorkbook.Worksheets("VarTab").Cells(iRow + 1, 3)) Then
                addrOffset2 = a
                Else
                addrOffset2 = Replace(TagType_array2(1), "DBX", "")
                End If
                
                If addrBit <> addrBit2 Then
                Do Until addrBit = addrBit2
                res2 = daveReadBytes(dc, daveDB, dbnum, addrOffset, 1, 0)
                bfbyte = daveGetU8(dc)                                         'Copy and convert a value of 8 bit, signed or unsigned from internal buffer
                bitStat = bfbyte And 2 ^ addrBit2
                If bitStat > 0 Then 'Convert byte to bit
                bitStat = 1
                Else
                bitStat = 0
                End If
                
                bitStat = bitStat * 2 ^ addrBit2                                 'Shift bit on the byte array
                tryone = bitStat + tryone
                addrBit2 = addrBit2 + 1
                
                Loop
                Else
                End If
                If addrOffset <> addrOffset2 And addrBit2 <> 7 Then
                value = ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 5)

                If value = True Then                                                'Convert string to bit
                value = 1
                Else
                value = 0
                End If
                    
                bitStat = value * 2 ^ addrBit                                       'Shift bit on the byte array
                tryone = bitStat + tryone
                addrBit2 = addrBit2 + 1
                
                Do Until addrBit2 > 7
                res2 = daveReadBytes(dc, daveDB, dbnum, addrOffset, 1, 0)
                bfbyte = daveGetU8(dc)                                             'Copy and convert a value of 8 bit, signed or unsigned from internal buffer
                bitStat = bfbyte And 2 ^ addrBit2
                If bitStat > 0 Then  'Convert byte to bit
                 bitStat = 1
                Else
                 bitStat = 0
                End If
                
                bitStat = bitStat * 2 ^ addrBit2                                      ' Shift bit on the byte array
                tryone = bitStat + tryone
                addrBit2 = addrBit2 + 1
                Loop
                Else
                End If
                
                
                value = ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 5)

                If value = True Then                                                'Convert string to bit
                value = 1
                Else
                value = 0
                End If
                    
                bitStat = value * 2 ^ addrBit                                       'Shift bit on the byte array
                tryone = bitStat + tryone
                addrBit2 = addrBit2 + 1
                
                If addrBit2 > 7 Then
                res2 = davePut8(buffer, tryone)                                     'Copy and convert a value of 8 bit, signed or unsigned into a buffer
                res2 = daveWriteBytes(dc, daveDB, dbnum, addrOffset, 1, buffer)     'Write a value or a block of values to PLC.
                addrBit2 = 0
                tryone = 0
                Else
                End If
                 
 '--------------------------------Write input data--------------------------------------
            ElseIf InStr(TagType_array(0), "I") > 0 Then
                addrOffset = Replace(TagType_array(0), "I", "")                             'Delete unnecessary Prefix in address
                addrBit = TagType_array(1)
                                
                If Not IsEmpty(ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 5)) Then     'check if cell is not empty
                    value = ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 5)
                    If value = True Then
                        res2 = daveSetBit(dc, daveInputs, areaNum, addrOffset, addrBit)  'set bit
                    ElseIf value = False Then
                        res2 = daveClrBit(dc, daveInputs, areaNum, addrOffset, addrBit)  'reset bit
                    End If
                End If
                
'--------------------------------Write output data--------------------------------------
            ElseIf InStr(TagType_array(0), "Q") > 0 Then
                addrOffset = Replace(TagType_array(0), "Q", "")                             'Delete unnecessary Prefix in address
                addrBit = TagType_array(1)
                                
                If Not IsEmpty(ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 5)) Then     'check if cell is not empty
                    value = ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 5)
                    If value = True Then
                        res2 = daveSetBit(dc, daveOutputs, areaNum, addrOffset, addrBit)  'set bit
                    ElseIf value = False Then
                        res2 = daveClrBit(dc, daveOutputs, areaNum, addrOffset, addrBit)  'reset bit
                    End If
                End If
                
 '--------------------------------Write dint data--------------------------------------
            Else
                If InStr(TagType_array(1), "DBDW") > 0 Then
                    addrOffset = Replace(TagType_array(1), "DBDW", "")
                    value = ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 5)
                    
                    res2 = davePut32(buffer3, value)                              'Copy and convert a value of 32 bit, signed or unsigned into a buffer
                    res2 = daveWriteBytes(dc, daveDB, dbnum, addrOffset, 4, buffer3) 'Write a value or a block of values to PLC.
                
'--------------------------------Write real data--------------------------------------
                ElseIf InStr(TagType_array(1), "DBD") > 0 Then
                    addrOffset = Replace(TagType_array(1), "DBD", "")
                    value = ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 5)
                    
                    res2 = davePutFloat(buffer1, value)                              'Copy and convert a value of 32 bit, signed or unsigned into a buffer
                    res2 = daveWriteBytes(dc, daveDB, dbnum, addrOffset, 4, buffer1) 'Write a value or a block of values to PLC.

'--------------------------------Write word data--------------------------------------
                ElseIf InStr(TagType_array(1), "DBW") > 0 Then
                    addrOffset = Replace(TagType_array(1), "DBW", "")
                    value = ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 5)
                    
                    res2 = davePut16(buffer2, value)                                 'Copy and convert a value of 16 bit, signed or unsigned into a buffer
                    res2 = daveWriteBytes(dc, daveDB, dbnum, addrOffset, 2, buffer2) 'Write a value or a block of values to PLC.

'--------------------------------Write byte data--------------------------------------
                ElseIf InStr(TagType_array(1), "DBB") > 0 Then
                    addrOffset = Replace(TagType_array(1), "DBB", "")
                    
                    value = ActiveWorkbook.Worksheets("VarTab").Cells(iRow, 5)
                    
                    res2 = davePut8(buffer, value)                                  'Copy and convert a value of 8 bit, signed or unsigned into a buffer
                    res2 = daveWriteBytes(dc, daveDB, dbnum, addrOffset, 1, buffer) 'Write a value or a block of values to PLC.

                End If
            End If
            iRow = iRow + 1
        Loop
    Else
        MsgBox "No route PLC, check connection and settings"
    End If
        Call cleanUp(ph, di, dc)
End Sub
