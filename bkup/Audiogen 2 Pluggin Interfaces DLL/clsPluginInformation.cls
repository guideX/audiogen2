VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPluginInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get PluginName(ByVal locale As String) As String
Attribute PluginName.VB_Description = "Gets the name of this plugin."

End Property

Public Property Get PluginDescription(ByVal locale As String) As String
Attribute PluginDescription.VB_Description = "Gets the description of this plugin."

End Property

Public Property Get PluginAcknowledgements(ByVal locale As String) As String
Attribute PluginAcknowledgements.VB_Description = "Gets any acknowledgements which should be displayed for this plugin."

End Property

Public Property Get PluginAuthor(ByVal locale As String) As String
Attribute PluginAuthor.VB_Description = "Gets the author of this plugin."

End Property

Public Property Get PluginWebsite(ByVal locale As String) As String
Attribute PluginWebsite.VB_Description = "Gets the website for this plugin."

End Property

Public Property Get Configuration() As clsPluginConfig
Attribute Configuration.VB_Description = "Gets the IPlugConfig interface for this plugin.  If the plugin does not have any configuration then return Nothing."

End Property
