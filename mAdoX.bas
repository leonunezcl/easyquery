Attribute VB_Name = "mAdo"
Option Explicit

Public Function TipoDeCampo(ByVal Tipo As DataTypeEnum) As String

    Dim ret As String
        
    If Tipo = adBigInt Then ret = "BigInt"
    If Tipo = adBinary Then ret = "Binary"
    If Tipo = adBoolean Then ret = "Boolean"
    If Tipo = adBSTR Then ret = "BSTR"
    If Tipo = adChar Then ret = "Char"
    If Tipo = adCurrency Then ret = "Currency"
    If Tipo = adDate Then ret = "Date"
    If Tipo = adDBDate Then ret = "Date"
    If Tipo = adDBTime Then ret = "Time"
    If Tipo = adDBTimeStamp Then ret = "TimeStamp"
    If Tipo = adDecimal Then ret = "Decimal"
    If Tipo = adDouble Then ret = "Double"
    If Tipo = adEmpty Then ret = "Empty"
    If Tipo = adError Then ret = "Error"
    If Tipo = adGUID Then ret = "GUID"
    If Tipo = adIDispatch Then ret = "IDispatch"
    If Tipo = adInteger Then ret = "Integer"
    If Tipo = adIUnknown Then ret = "Unknown"
    If Tipo = adLongVarBinary Then ret = "LongVarBinary"
    If Tipo = adLongVarChar Then ret = "LongVarChar"
    If Tipo = adLongVarWChar Then ret = "LongVarWChar"
    If Tipo = adNumeric Then ret = "Numeric"
    If Tipo = adSingle Then ret = "Single"
    If Tipo = adSmallInt Then ret = "SmallInt"
    If Tipo = adTinyInt Then ret = "TinyInt"
    If Tipo = adUnsignedBigInt Then ret = "UnsignedBigInt"
    If Tipo = adUnsignedInt Then ret = "UnsignedInt"
    If Tipo = adUnsignedSmallInt Then ret = "UnsignedSmallInt"
    If Tipo = adUnsignedTinyInt Then ret = "UnsignedTinyInt"
    If Tipo = adUserDefined Then ret = "UserDefined"
    If Tipo = adVarBinary Then ret = "VarBinary"
    If Tipo = adVarChar Then ret = "VarChar"
    If Tipo = adVariant Then ret = "Variant"
    If Tipo = adVarWChar Then ret = "VarWChar"
    If Tipo = adWChar Then ret = "WChar"
        
    TipoDeCampo = ret
    
End Function

