Attribute VB_Name = "modGeneral"
Option Explicit

Sub Main()
     Const SourceConfigFile As String = "E:\Projects\Builder\ZDoom_Doom.cfg"
     Const TargetConfigFile As String = "E:\Projects\Builder\ZDoom_DoomHexen.cfg"
     Const OutputConfigFile As String = "E:\Projects\Builder\output.cfg"
     Dim source As New clsConfiguration
     Dim target As New clsConfiguration
     Dim output As New clsConfiguration
     Dim sourcethings As Dictionary
     Dim outputthings As Dictionary
     Dim catkeys As Variant
     Dim tcat As Dictionary
     Dim scat As Dictionary
     Dim c As Long
     Dim t As Long
     Dim thingkeys As Variant
     Dim tthing As Dictionary
     Dim sthing As Dictionary
     
     'Open configurations
     source.LoadConfiguration SourceConfigFile
     target.LoadConfiguration TargetConfigFile
     
     'Read
     Set sourcethings = source.ReadSetting("thingtypes", New Dictionary, True)
     Set outputthings = target.ReadSetting("thingtypes", New Dictionary, False)
     
     'Go for all categories in target
     catkeys = outputthings.Keys
     For c = LBound(catkeys) To UBound(catkeys)
          
          'Get categorie
          Set tcat = outputthings(catkeys(c))
          
          'Does this categorie exist in source?
          If (sourcethings.Exists(catkeys(c))) Then
               
               'Get categorie
               Set scat = sourcethings(catkeys(c))
               
               'Does the source have a key, but the target hasnt?
               'Then copy this key-value over to target
               If (scat.Exists("width") = True) And (tcat.Exists("width") = False) Then tcat.Add "width", scat("width")
               If (scat.Exists("height") = True) And (tcat.Exists("height") = False) Then tcat.Add "height", scat("height")
               If (scat.Exists("hangs") = True) And (tcat.Exists("hangs") = False) Then tcat.Add "hangs", scat("hangs")
               If (scat.Exists("blocking") = True) And (tcat.Exists("blocking") = False) Then tcat.Add "blocking", scat("blocking")
               If (scat.Exists("error") = True) And (tcat.Exists("error") = False) Then tcat.Add "error", scat("error")
               
               'Go for all things in categorie
               thingkeys = tcat.Keys
               For t = LBound(thingkeys) To UBound(thingkeys)
                    
                    'Check if this is a thing
                    If (IsNumeric(thingkeys(t))) Then
                         
                         'If the thing is only a string, then expand
                         If (Not IsObject(tcat(thingkeys(t)))) Then
                              
                              'Create thing
                              Set tthing = New Dictionary
                              tthing.Add "title", CStr(tcat(thingkeys(t)))
                              Set tcat.Item(thingkeys(t)) = tthing
                         Else
                              
                              'Get the thing
                              Set tthing = tcat(thingkeys(t))
                         End If
                         
                         'Does this thing exist in source?
                         If (scat.Exists(thingkeys(t))) Then
                              
                              'If the thing is only a string, then expand
                              If (Not IsObject(scat(thingkeys(t)))) Then
                                   
                                   'Create thing
                                   Set sthing = New Dictionary
                                   sthing.Add "title", CStr(scat(thingkeys(t)))
                              Else
                                   
                                   'Get the thing
                                   Set sthing = scat(thingkeys(t))
                              End If
                              
                              'Does the source have a key, but the target hasnt?
                              'Then copy this key-value over to target
                              If (sthing.Exists("width") = True) And (tthing.Exists("width") = False) Then tthing.Add "width", sthing("width")
                              If (sthing.Exists("height") = True) And (tthing.Exists("height") = False) Then tthing.Add "height", sthing("height")
                              If (sthing.Exists("sprite") = True) And (tthing.Exists("sprite") = False) Then tthing.Add "sprite", sthing("sprite")
                              If (sthing.Exists("hangs") = True) And (tthing.Exists("hangs") = False) Then tthing.Add "hangs", sthing("hangs")
                              If (sthing.Exists("blocking") = True) And (tthing.Exists("blocking") = False) Then tthing.Add "blocking", sthing("blocking")
                              If (sthing.Exists("error") = True) And (tthing.Exists("error") = False) Then tthing.Add "error", sthing("error")
                         End If
                    End If
               Next t
          End If
     Next c
     
     'Write
     output.WriteSetting "thingtypes", outputthings, True
     output.SaveConfiguration OutputConfigFile
End Sub


