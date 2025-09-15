[_ISTool]
EnableISX=false

[Setup]
OutputDir=E:\Projects\Builder\Package
SourceDir=E:\Projects\Builder\Package\Support
OutputBaseFilename=builder_setup
Compression=zip/9
AppName=Doom Builder
AppVerName=Doom Builder
AppMutex=DoomBuilder
AllowNoIcons=true
DefaultGroupName=CodeImp

AlwaysShowComponentsList=false
AppPublisher=CodeImp
AppPublisherURL=http://www.codeimp.com
AppSupportURL=http://www.doombuilder.com
AppUpdatesURL=http://www.doombuilder.com
EnableDirDoesntExistWarning=false
DirExistsWarning=auto
MinVersion=4.1.1998,5.0.2195
InfoBeforeFile=E:\Projects\Singleline Disclaimer.txt
ChangesAssociations=false
BackColor=clMaroon
WizardImageBackColor=$000080
UninstallDisplayIcon={app}\Builder.exe
DefaultDirName={pf}\Doom Builder\
ShowLanguageDialog=yes

[Dirs]

[Files]
Source: COMCAT.DLL; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall noregerror
Source: ASYCFILT.DLL; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall noregerror
Source: Msvbvm60.dll; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall noregerror
Source: OLEAUT32.DLL; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall
Source: OLEPRO32.DLL; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall
Source: scrrun.dll; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall
Source: STDOLE2.TLB; DestDir: {sys}; Flags: restartreplace sharedfile uninsneveruninstall regtypelib noregerror
Source: mscomctl.ocx; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall overwritereadonly
Source: Builder.exe; DestDir: {app}; Flags: ignoreversion
Source: Boom.cfg; DestDir: {app}; Flags: ignoreversion
Source: Builder.cfg; DestDir: {app}; Flags: onlyifdoesntexist
Source: Builder.dll; DestDir: {app}; Flags: ignoreversion
Source: Crosshair.bmp; DestDir: {app}; Flags: ignoreversion
Source: Doom.cfg; DestDir: {app}; Flags: ignoreversion
Source: Doom2.cfg; DestDir: {app}; Flags: ignoreversion
Source: Font.fnt; DestDir: {app}; Flags: ignoreversion
Source: GNU_GPL.txt; DestDir: {app}; Flags: ignoreversion
Source: Heretic.cfg; DestDir: {app}; Flags: ignoreversion
Source: Shortcuts.cfg; DestDir: {app}; Flags: ignoreversion
Source: UltDoom.cfg; DestDir: {app}; Flags: ignoreversion
Source: acc.exe; DestDir: {app}; Flags: ignoreversion
Source: bsp-w32.exe; DestDir: {app}; Flags: ignoreversion
Source: common.acs; DestDir: {app}; Flags: ignoreversion
Source: deacc.exe; DestDir: {app}; Flags: ignoreversion
Source: defs.acs; DestDir: {app}; Flags: ignoreversion
Source: Eternity.cfg; DestDir: {app}; Flags: ignoreversion
Source: Hexen.cfg; DestDir: {app}; Flags: ignoreversion
Source: Legacy.cfg; DestDir: {app}; Flags: ignoreversion
Source: Missing.bmp; DestDir: {app}; Flags: ignoreversion
Source: specials.acs; DestDir: {app}; Flags: ignoreversion
Source: Unknown.bmp; DestDir: {app}; Flags: ignoreversion
Source: wvars.acs; DestDir: {app}; Flags: ignoreversion
Source: zcommon.acs; DestDir: {app}; Flags: ignoreversion
Source: zdefs.acs; DestDir: {app}; Flags: ignoreversion
Source: ZDoom_Doom.cfg; DestDir: {app}; Flags: ignoreversion
Source: ZDoom_DoomHexen.cfg; DestDir: {app}; Flags: ignoreversion
Source: ZDoom_Hexen.cfg; DestDir: {app}; Flags: ignoreversion
Source: zspecial.acs; DestDir: {app}; Flags: ignoreversion
Source: zwvars.acs; DestDir: {app}; Flags: ignoreversion
Source: fbase6.txt; DestDir: {app}; Flags: ignoreversion
Source: fbase6.wad; DestDir: {app}; Flags: ignoreversion
Source: jDoom.cfg; DestDir: {app}; Flags: ignoreversion
Source: Skulltag_Doom.cfg; DestDir: {app}; Flags: ignoreversion
Source: Skulltag_DoomHexen.cfg; DestDir: {app}; Flags: ignoreversion
Source: ACS.cfg; DestDir: {app}; Flags: ignoreversion
Source: FS.cfg; DestDir: {app}; Flags: ignoreversion
Source: zdbsp.exe; DestDir: {app}; Flags: ignoreversion
Source: ZenNode.exe; DestDir: {app}; Flags: ignoreversion
Source: Parameters.cfg; DestDir: {app}; Flags: onlyifdoesntexist
Source: Font.tga; DestDir: {app}; Flags: ignoreversion
Source: GDIPlus.dll; DestDir: {app}; Flags: ignoreversion
Source: GDIPlus.tlb; DestDir: {sys}; Flags: restartreplace sharedfile uninsneveruninstall regtypelib noregerror
Source: Strife.cfg; DestDir: {app}; Flags: ignoreversion
Source: cmcs21.dll; DestDir: {sys}; Flags: restartreplace sharedfile uninsneveruninstall overwritereadonly
Source: cmcs21.ocx; DestDir: {sys}; Flags: regserver restartreplace sharedfile uninsneveruninstall overwritereadonly
Source: Risen3D.cfg; DestDir: {app}; Flags: ignoreversion
Source: DED.cfg; DestDir: {app}; Flags: ignoreversion
Source: Dehacked.cfg; DestDir: {app}; Flags: ignoreversion
Source: Thingarrow.tga; DestDir: {app}; Flags: ignoreversion
Source: Thingbox.tga; DestDir: {app}; Flags: ignoreversion
Source: ZDoom_StrifeHexen.cfg; DestDir: {app}; Flags: ignoreversion
Source: ZDoom_HereticHexen.cfg; DestDir: {app}; Flags: ignoreversion

[Icons]

Name: {group}\Doom Builder; Filename: {app}\Builder.exe; WorkingDir: {app}; IconFilename: {app}\Builder.exe; Comment: Doom Builder - The cornerstone for every map author; IconIndex: 0; Flags: createonlyiffileexists
Name: {group}\CodeImp website; Filename: http://www.codeimp.com/; IconIndex: 0
Name: {group}\Doom Builder website; Filename: http://www.doombuilder.com/; IconIndex: 0

[UninstallDelete]

[Run]
Filename: {app}\Builder.exe; WorkingDir: {app}; Flags: unchecked skipifdoesntexist postinstall skipifsilent; Description: Launch Doom Builder now
