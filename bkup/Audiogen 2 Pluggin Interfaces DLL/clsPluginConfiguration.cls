VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPluginConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get ParameterCount() As Long
Attribute ParameterCount.VB_Description = "Gets the number of parameters for this plugin."

End Property

Public Property Get ParameterName(ByVal locale As String, ByVal paramNumber As Long) As String
Attribute ParameterName.VB_Description = "Gets the name of the parameter with the specified index."

End Property

Public Property Get ParameterValueCount(ByVal paramNumber As Long) As Long
Attribute ParameterValueCount.VB_Description = "Gets the number of values for the parameter with the specified index."

End Property

Public Property Get ParameterValueName(ByVal locale As String, ByVal paramNumber As Long, ByVal valueNumber As Long) As String
Attribute ParameterValueName.VB_Description = "Gets the name of the value for the specified parameter value."

End Property

Public Property Get ParameterDefaultValue(ByVal valueNumber As Long) As Long
Attribute ParameterDefaultValue.VB_Description = "Gets the default value index for the specified parameter index."

End Property

Public Sub SetParameterValue(ByVal locale As String, ByVal paramNumber As Long, ByVal valueNumber As Long, ByRef changedParameterIndexes As Variant)
Attribute SetParameterValue.VB_Description = "Called by the client when a parameter value is changed."

End Sub
