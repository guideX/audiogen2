VERSION 5.00
Begin VB.UserControl ctlMusicPlayer 
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   735
   Picture         =   "ctlMusicPlayer.ctx":0000
   ScaleHeight     =   645
   ScaleWidth      =   735
   Windowless      =   -1  'True
End
Attribute VB_Name = "ctlMusicPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private lInitialized As Boolean
Private lSongHandle As Long
Private lStreamHandle As Long
Private lStreamChannel As Long
Private Const FMOD_VERSION = 3.74
Enum FMOD_ERRORS
    FMOD_ERR_NONE             ' No errors
    FMOD_ERR_BUSY             ' Cannot call this command after FSOUND_Init.  Call FSOUND_Close first.
    FMOD_ERR_UNINITIALIZED    ' This command failed because FSOUND_Init was not called
    FMOD_ERR_INIT             ' Error initializing output device.
    FMOD_ERR_ALLOCATED        ' Error initializing output device, but more specifically, the output device is already in use and cannot be reused.
    FMOD_ERR_PLAY             ' Playing the sound failed.
    FMOD_ERR_OUTPUT_FORMAT    ' Soundcard does not support the features needed for this soundsystem (16bit stereo output)
    FMOD_ERR_COOPERATIVELEVEL ' Error setting cooperative level for hardware.
    FMOD_ERR_CREATEBUFFER     ' Error creating hardware sound buffer.
    FMOD_ERR_FILE_NOTFOUND    ' File not found
    FMOD_ERR_FILE_FORMAT      ' Unknown file format
    FMOD_ERR_FILE_BAD         ' Error loading file
    FMOD_ERR_MEMORY           ' Not enough memory
    FMOD_ERR_VERSION          ' The version number of this file format is not supported
    FMOD_ERR_INVALID_PARAM    ' An invalid parameter was passed to this function
    FMOD_ERR_NO_EAX           ' Tried to use an EAX command on a non EAX enabled channel or output.
    FMOD_ERR_CHANNEL_ALLOC    ' Failed to allocate a new channel
    FMOD_ERR_RECORD           ' Recording is not supported on this machine
    FMOD_ERR_MEDIAPLAYER      ' Windows Media Player not installed so cannot play wma or use internet streaming.
    FMOD_ERR_CDDEVICE         ' An error occured trying to open the specified CD device
End Enum
Enum FSOUND_OUTPUTTYPES
    FSOUND_OUTPUT_NOSOUND   ' NoSound driver, all calls to this succeed but do nothing.
    FSOUND_OUTPUT_WINMM     ' Windows Multimedia driver.
    FSOUND_OUTPUT_DSOUND    ' DirectSound driver.  You need this to get EAX2 or EAX3 support, or FX api support.
    FSOUND_OUTPUT_A3D       ' A3D driver.
    FSOUND_OUTPUT_OSS       ' Linux/Unix OSS (Open Sound System) driver, i.e. the kernel sound drivers.
    FSOUND_OUTPUT_ESD       ' Linux/Unix ESD (Enlightment Sound Daemon) driver.
    FSOUND_OUTPUT_ALSA      ' Linux Alsa driver.
    FSOUND_OUTPUT_ASIO      ' Low latency ASIO driver
    FSOUND_OUTPUT_XBOX      ' Xbox driver
    FSOUND_OUTPUT_PS2       ' PlayStation 2 driver
    FSOUND_OUTPUT_MAC       ' Mac SoundMager driver
    FSOUND_OUTPUT_GC        ' Gamecube driver
    FSOUND_OUTPUT_NOSOUND_NONREALTIME  ' This is the same as nosound, but the sound generation is driven by FSOUND_Update
End Enum
Enum FSOUND_MIXERTYPES
    FSOUND_MIXER_AUTODETECT         ' CE/PS2 Only - Non interpolating/low quality mixer
    FSOUND_MIXER_BLENDMODE          ' removed / obsolete.
    FSOUND_MIXER_MMXP5              ' removed / obsolete.
    FSOUND_MIXER_MMXP6              ' removed / obsolete.
    FSOUND_MIXER_QUALITY_AUTODETECT ' All platforms - Autodetect the fastest quality mixer based on your cpu.
    FSOUND_MIXER_QUALITY_FPU        ' Win32/Linux only - Interpolating/volume ramping FPU mixer.
    FSOUND_MIXER_QUALITY_MMXP5      ' Win32/Linux only - Interpolating/volume ramping FPU mixer.
    FSOUND_MIXER_QUALITY_MMXP6      ' Win32/Linux only - Interpolating/volume ramping ppro+ MMX mixer.
    FSOUND_MIXER_MONO               ' CE/PS2 only - MONO non interpolating/low quality mixer. For speed
    FSOUND_MIXER_QUALITY_MONO       ' CE/PS2 only - MONO Interpolating mixer.  For speed
End Enum
Private Enum FMUSIC_TYPES
    FMUSIC_TYPE_NONE
    FMUSIC_TYPE_MOD         'Protracker / Fasttracker
    FMUSIC_TYPE_S3M         'ScreamTracker 3
    FMUSIC_TYPE_XM          'FastTracker 2
    FMUSIC_TYPE_IT          'Impulse Tracker.
    FMUSIC_TYPE_MIDI        'MIDI file
    FMUSIC_TYPE_FSB         'FMOD Sample Bank file
End Enum
Private Enum FSOUND_DSP_PRIORITIES
    FSOUND_DSP_DEFAULTPRIORITY_CLEARUNIT = 0           'DSP CLEAR unit - done first
    FSOUND_DSP_DEFAULTPRIORITY_SFXUNIT = 100           'DSP SFX unit - done second
    FSOUND_DSP_DEFAULTPRIORITY_MUSICUNIT = 200         'DSP MUSIC unit - done third
    FSOUND_DSP_DEFAULTPRIORITY_USER = 300              'User priority, use this as reference for your own dsp units
    FSOUND_DSP_DEFAULTPRIORITY_FFTUNIT = 900           'This reads data for FSOUND_DSP_GetSpectrum, so it comes after user units
    FSOUND_DSP_DEFAULTPRIORITY_CLIPANDCOPYUNIT = 1000  'DSP CLIP AND COPY unit - last
End Enum
Private Enum FSOUND_CAPS
    FSOUND_CAPS_HARDWARE = &H1       ' This driver supports hardware accelerated 3d sound.
    FSOUND_CAPS_EAX2 = &H2           ' This driver supports EAX 2 reverb
    FSOUND_CAPS_EAX3 = &H10          ' This driver supports EAX 3 reverb
End Enum
Private Enum FSOUND_MODES
    FSOUND_LOOP_OFF = &H1             ' For non looping samples.
    FSOUND_LOOP_NORMAL = &H2          ' For forward looping samples.
    FSOUND_LOOP_BIDI = &H4            ' For bidirectional looping samples.  (no effect if in hardware).
    FSOUND_8BITS = &H8                ' For 8 bit samples.
    FSOUND_16BITS = &H10              ' For 16 bit samples.
    FSOUND_MONO = &H20                ' For mono samples.
    FSOUND_STEREO = &H40              ' For stereo samples.
    FSOUND_UNSIGNED = &H80            ' For source data containing unsigned samples.
    FSOUND_SIGNED = &H100             ' For source data containing signed data.
    FSOUND_DELTA = &H200              ' For source data stored as delta values.
    FSOUND_IT214 = &H400              ' For source data stored using IT214 compression.
    FSOUND_IT215 = &H800              ' For source data stored using IT215 compression.
    FSOUND_HW3D = &H1000              ' Attempts to make samples use 3d hardware acceleration. (if the card supports it)
    FSOUND_2D = &H2000                ' Ignores any 3d processing.  overrides FSOUND_HW3D.  Located in software.
    FSOUND_STREAMABLE = &H4000        ' For realtime streamable samples.  If you dont supply this sound may come out corrupted.
    FSOUND_LOADMEMORY = &H8000        ' For FSOUND_Sample_Load - name will be interpreted as a pointer to data
    FSOUND_LOADRAW = &H10000          ' For FSOUND_Sample_Load/FSOUND_Stream_Open - will ignore file format and treat as raw pcm.
    FSOUND_MPEGACCURATE = &H20000     ' For FSOUND_Stream_Open - scans MP2/MP3 (VBR also) for accurate FSOUND_Stream_GetLengthMs/FSOUND_Stream_SetTime.
    FSOUND_FORCEMONO = &H40000        ' For forcing stereo streams and samples to be mono - needed with FSOUND_HW3D - incurs speed hit
    FSOUND_HW2D = &H80000             ' 2d hardware sounds.  allows hardware specific effects
    FSOUND_ENABLEFX = &H100000        ' Allows DX8 FX to be played back on a sound.  Requires DirectX 8 - Note these sounds cant be played more than once, or have a changing frequency
    FSOUND_MPEGHALFRATE = &H200000    ' For FMODCE only - decodes mpeg streams using a lower quality decode, but faster execution
    FSOUND_XADPCM = &H400000          ' For XBOX only - Describes a user sample that its contents are compressed as XADPCM
    FSOUND_VAG = &H800000             ' For PS2 only - Describes a user sample that its contents are compressed as Sony VAG format.
    FSOUND_NONBLOCKING = &H1000000    ' For FSOUND_Stream_Open - Causes stream to open in the background and not block the foreground app - stream plays only when ready.
    FSOUND_GCADPCM = &H2000000        ' For Gamecube only - Contents are compressed as Gamecube DSP-ADPCM format
    FSOUND_MULTICHANNEL = &H4000000   ' For PS2 only - Contents are interleaved into a multi-channel (more than stereo) format
    FSOUND_USECORE0 = &H8000000       ' For PS2 only - Sample/Stream is forced to use hardware voices 00-23
    FSOUND_USECORE1 = &H10000000      ' For PS2 only - Sample/Stream is forced to use hardware voices 24-47
    FSOUND_LOADMEMORYIOP = &H20000000 ' For PS2 only - "name" will be interpreted as a pointer to data for streaming and samples.  The address provided will be an IOP address
    FSOUND_STREAM_NET = &H80000000    ' Specifies an internet stream
    FSOUND_NORMAL = FSOUND_16BITS Or FSOUND_SIGNED Or FSOUND_MONO
End Enum
Private Enum FSOUND_CDPLAYMODES
    FSOUND_CD_PLAYCONTINUOUS
    FSOUND_CD_PLAYONCE
    FSOUND_CD_PLAYLOOPED
    FSOUND_CD_PLAYRANDOM
End Enum
Private Enum FSOUND_CHANNELSAMPLEMODE
    FSOUND_FREE = -1                 ' definition for dynamically allocated channel or sample
    FSOUND_UNMANAGED = -2            ' definition for allocating a sample that is NOT managed by fsound
    FSOUND_ALL = -3                  ' for a channel index or sample index, this flag affects ALL channels or samples available!  Not supported by all functions.
    FSOUND_STEREOPAN = -1            ' definition for full middle stereo volume on both channels
    FSOUND_SYSTEMCHANNEL = -1000     ' special channel ID for channel based functions that want to alter the global FSOUND software mixing output channel
    FSOUND_SYSTEMSAMPLE = -1000      ' special sample ID for all sample based functions that want to alter the global FSOUND software mixing output sample
End Enum
Private Type FSOUND_REVERB_PROPERTIES
    Environment         As Long     ' 0       25     0       sets all listener properties
    EnvSize             As Single   ' 1.0     100.0  7.5     environment size in meters
    EnvDiffusion        As Single   ' 0.0     1.0    1.0     environment diffusion
    Room                As Long     ' -10000  0      -1000   room effect level (at mid frequencies)
    RoomHF              As Long     ' -10000  0      -100    relative room effect level at high frequencies
    RoomLF              As Long     ' -10000  0      0       relative room effect level at low frequencies
    DecayTime           As Single   ' 0.1     20.0   1.49    reverberation decay time at mid frequencies
    DecayHFRatio        As Single   ' 0.1     2.0    0.83    high-frequency to mid-frequency decay time ratio
    DecayLFRatio        As Single   ' 0.1     2.0    1.0     low-frequency to mid-frequency decay time ratio
    Reflections         As Long     ' -10000  1000   -2602   early reflections level relative to room effect
    ReflectionsDelay    As Single   ' 0.0     0.3    0.007   initial reflection delay time
    ReflectionsPan(3)   As Single   '                0,0,0   early reflections panning vector
    Reverb              As Long     ' -1000   2000   200     late reverberation level relative to room effect
    ReverbDelay         As Single   ' 0.0     0.1    0.011   late reverberation delay time relative to initial reflection
    ReverbPan(3)        As Single   '                0,0,0   late reverberation panning vector
    EchoTime            As Single   ' .075    0.25   0.25    echo time
    EchoDepth           As Single   ' 0.0     1.0    0.0     echo depth
    ModulationTime      As Single   ' 0.04    4.0    0.25    modulation time
    ModulationDepth     As Single   ' 0.0     1.0    0.0     modulation depth
    AirAbsorptionHF     As Single   ' -100    0.0    -5.0    change in level per meter at high frequencies
    HFReference         As Single   ' 1000.0  20000  5000.0  reference high frequency (hz)
    LFReference         As Single   ' 20.0    1000.0 250.0   reference low frequency (hz)
    RoomRolloffFactor   As Single   ' 0.0     10.0   0.0     like FSOUND_3D_SetRolloffFactor but for room effect
    Diffusion           As Single   ' 0.0     100.0  100.0   Value that controls the echo density in the late reverberation decay. (xbox only)
    Density             As Single   ' 0.0     100.0  100.0   Value that controls the modal density in the late reverberation decay (xbox only)
    flags               As Long     '                        modifies the behavior of above properties
End Type
Private Enum FSOUND_REVERB_PROPERTYFLAGS
    FSOUND_REVERBFLAGS_DECAYTIMESCALE = &H1          ' EnvironmentSize affects reverberation decay time
    FSOUND_REVERBFLAGS_REFLECTIONSSCALE = &H2        ' EnvironmentSize affects reflection level
    FSOUND_REVERBFLAGS_REFLECTIONSDELAYSCALE = &H4   ' EnvironmentSize affects initial reflection delay time
    FSOUND_REVERBFLAGS_REVERBSCALE = &H8             ' EnvironmentSize affects reflections level
    FSOUND_REVERBFLAGS_REVERBDELAYSCALE = &H10       ' EnvironmentSize affects late reverberation delay time
    FSOUND_REVERBFLAGS_DECAYHFLIMIT = &H20           ' AirAbsorptionHF affects DecayHFRatio
    FSOUND_REVERBFLAGS_ECHOTIMESCALE = &H40          ' EnvironmentSize affects echo time
    FSOUND_REVERBFLAGS_MODULATIONTIMESCALE = &H80    ' EnvironmentSize affects modulation time
    FSOUND_REVERB_FLAGS_CORE0 = &H100                ' PS2 Only - Reverb is applied to CORE0 (hw voices 0-23)
    FSOUND_REVERB_FLAGS_CORE1 = &H200                ' PS2 Only - Reverb is applied to CORE1 (hw voices 24-47)
    FSOUND_REVERBFLAGS_DEFAULT = FSOUND_REVERBFLAGS_DECAYTIMESCALE Or FSOUND_REVERBFLAGS_REFLECTIONSSCALE Or FSOUND_REVERBFLAGS_REFLECTIONSDELAYSCALE Or FSOUND_REVERBFLAGS_REVERBSCALE Or FSOUND_REVERBFLAGS_REVERBDELAYSCALE Or FSOUND_REVERBFLAGS_DECAYHFLIMIT Or FSOUND_REVERB_FLAGS_CORE0 Or FSOUND_REVERB_FLAGS_CORE1
End Enum
Private Type FSOUND_REVERB_CHANNELPROPERTIES
    Direct               As Long     ' direct path level (at low and mid frequencies)
    DirectHF             As Long     ' relative direct path level at high frequencies
    Room                 As Long     ' room effect level (at low and mid frequencies)
    RoomHF               As Long     ' relative room effect level at high frequencies
    Obstruction          As Long     ' main obstruction control (attenuation at high frequencies)
    ObstructionLFRatio   As Single   ' obstruction low-frequency level re. main control
    Occlusion            As Long     ' main occlusion control (attenuation at high frequencies)
    OcclustionLFRatio    As Single   ' occlusion low-frequency level re. main control
    OcclusionRoomRatio   As Single   ' relative occlusion control for room effect
    OcclusionDirectRatio As Single   ' relative occlusion control for direct path
    Exclusion            As Long     ' main exlusion control (attenuation at high frequencies)
    ExclusionLFRatio     As Single   ' exclusion low-frequency level re. main control
    OutsideVolumeHF      As Long     ' outside sound cone level at high frequencies
    DopplerFactor        As Single   ' like DS3D flDopplerFactor but per source
    RolloffFactor        As Single   ' like DS3D flRolloffFactor but per source
    RoomRolloffFactor    As Single   ' like DS3D flRolloffFactor but for room effect
    AirAbsorptionFactor  As Single   ' multiplies AirAbsorptionHF member of FSOUND_REVERB_PROPERTIES
    flags                As Long     ' modifies the behavior of properties
End Type
Private Enum FSOUND_REVERB_CHANNELFLAGS
    FSOUND_REVERB_CHANNELFLAGS_DIRECTHFAUTO = &H1   ' Automatic setting of Direct due to distance from listener
    FSOUND_REVERB_CHANNELFLAGS_ROOMAUTO = &H2       ' Automatic setting of Room due to distance from listener
    FSOUND_REVERB_CHANNELFLAGS_ROOMHFAUTO = &H4     ' Automatic setting of RoomHF due to distance from listener
    FSOUND_REVERB_CHANNELFLAGS_DEFAULT = FSOUND_REVERB_CHANNELFLAGS_DIRECTHFAUTO Or FSOUND_REVERB_CHANNELFLAGS_ROOMAUTO Or FSOUND_REVERB_CHANNELFLAGS_ROOMHFAUTO
End Enum
Private Enum FSOUND_FX_MODES
    FSOUND_FX_CHORUS
    FSOUND_FX_COMPRESSOR
    FSOUND_FX_DISTORTION
    FSOUND_FX_ECHO
    FSOUND_FX_FLANGER
    FSOUND_FX_GARGLE
    FSOUND_FX_I3DL2REVERB
    FSOUND_FX_PARAMEQ
    FSOUND_FX_WAVES_REVERB
End Enum
Enum FSOUND_SPEAKERMODES
    FSOUND_SPEAKERMODE_DOLBYDIGITAL  ' The audio is played through a speaker arrangement of surround speakers with a subwoofer.
    FSOUND_SPEAKERMODE_HEADPHONE     ' The speakers are headphones.
    FSOUND_SPEAKERMODE_MONO          ' The speakers are monaural.
    FSOUND_SPEAKERMODE_QUAD          ' The speakers are quadraphonic.
    FSOUND_SPEAKERMODE_STEREO        ' The speakers are stereo (default value).
    FSOUND_SPEAKERMODE_SURROUND      ' The speakers are surround sound.
    FSOUND_SPEAKERMODE_DTS           ' The audio is played through a speaker arrangement of surround speakers with a subwoofer.
    FSOUND_SPEAKERMODE_PROLOGIC2     ' Dolby Prologic 2.  Playstation 2 and Gamecube only
End Enum
Private Enum FSOUND_INITMODES
    FSOUND_INIT_USEDEFAULTMIDISYNTH = &H1       'Causes MIDI playback to force software decoding.
    FSOUND_INIT_GLOBALFOCUS = &H2               'For DirectSound output - sound is not muted when window is out of focus.
    FSOUND_INIT_ENABLESYSTEMCHANNELFX = &H4     'For DirectSound output - Allows FSOUND_FX api to be used on global software mixer output!
    FSOUND_INIT_ACCURATEVULEVELS = &H8          'This latency adjusts FSOUND_GetCurrentLevels, but incurs a small cpu and memory hit
    FSOUND_INIT_PS2_DISABLECORE0REVERB = &H10   'PS2 only - Disable reverb on CORE 0 to regain SRAM
    FSOUND_INIT_PS2_DISABLECORE1REVERB = &H20   'PS2 only - Disable reverb on CORE 1 to regain SRAM
    FSOUND_INIT_PS2_SWAPDMACORES = &H40         'PS2 only - By default FMOD uses DMA CH0 for mixing, CH1 for uploads, this flag swaps them around
    FSOUND_INIT_DONTLATENCYADJUST = &H80        'Callbacks are not latency adjusted, and are called at mix time.  Also information functions are immediate
    FSOUND_INIT_GC_INITLIBS = &H100             'Gamecube only - Initializes GC audio libraries
    FSOUND_INIT_STREAM_FROM_MAIN_THREAD = &H200 'Turns off fmod streamer thread, and makes streaming update from FSOUND_Update called by the user
End Enum
Private Enum FSOUND_STREAM_NET_STATUS
    FSOUND_STREAM_NET_NOTCONNECTED         ' Stream hasn't connected yet
    FSOUND_STREAM_NET_CONNECTING           ' Stream is connecting to remote host
    FSOUND_STREAM_NET_BUFFERING            ' Stream is buffering data
    FSOUND_STREAM_NET_READY                ' Stream is ready to play
    FSOUND_STREAM_NET_ERROR                ' Stream has suffered a fatal error
End Enum
Private Enum FSOUND_TAGFIELD_TYPE
    FSOUND_TAGFIELD_VORBISCOMMENT = 0     ' A vorbis comment
    FSOUND_TAGFIELD_ID3V1                 ' Part of an ID3v1 tag
    FSOUND_TAGFIELD_ID3V2                 ' An ID3v2 frame
    FSOUND_TAGFIELD_SHOUTCAST             ' A SHOUTcast header line
    FSOUND_TAGFIELD_ICECAST               ' An Icecast header line
    FSOUND_TAGFIELD_ASF                   ' An Advanced Streaming Format header line
End Enum
Private Enum FSOUND_STATUS_FLAGS
    FSOUND_PROTOCOL_SHOUTCAST = &H1
    FSOUND_PROTOCOL_ICECAST = &H2
    FSOUND_PROTOCOL_HTTP = &H4
    FSOUND_FORMAT_MPEG = &H10000
    FSOUND_FORMAT_OGGVORBIS = &H20000
End Enum
Private Type FSOUND_TOC_TAG
    TagName(3)      As Byte         ' The string "TOC" (4th character is 0), just in case this structure is accidentally treated as a string.
    NumTracks       As Long         ' The number of tracks on the CD.
    min(99)         As Long         ' The start offset of each track in minutes.
    Sec(99)         As Long         ' The start offset of each track in seconds.
    Frame(99)       As Long         ' The start offset of each track in frames.
End Type
Private Declare Function FSOUND_SetOutput Lib "fmod.dll" Alias "_FSOUND_SetOutput@4" (ByVal outputtype As FSOUND_OUTPUTTYPES) As Byte
Private Declare Function FSOUND_SetDriver Lib "fmod.dll" Alias "_FSOUND_SetDriver@4" (ByVal driver As Long) As Byte
Private Declare Function FSOUND_SetMixer Lib "fmod.dll" Alias "_FSOUND_SetMixer@4" (ByVal mixer As FSOUND_MIXERTYPES) As Byte
Private Declare Function FSOUND_SetBufferSize Lib "fmod.dll" Alias "_FSOUND_SetBufferSize@4" (ByVal lenms As Long) As Byte
Private Declare Function FSOUND_SetHWND Lib "fmod.dll" Alias "_FSOUND_SetHWND@4" (ByVal hwnd As Long) As Byte
Private Declare Function FSOUND_SetMinHardwareChannels Lib "fmod.dll" Alias "_FSOUND_SetMinHardwareChannels@4" (ByVal min As Integer) As Byte
Private Declare Function FSOUND_SetMaxHardwareChannels Lib "fmod.dll" Alias "_FSOUND_SetMaxHardwareChannels@4" (ByVal min As Integer) As Byte
Private Declare Function FSOUND_SetMemorySystem Lib "fmod.dll" Alias "_FSOUND_SetMemorySystem@20" (ByVal pool As Long, ByVal poollen As Long, ByVal useralloc As Long, ByVal userrealloc As Long, ByVal userfree As Long) As Byte
Private Declare Function FSOUND_Init Lib "fmod.dll" Alias "_FSOUND_Init@12" (ByVal mixrate As Long, ByVal maxchannels As Long, ByVal flags As FSOUND_INITMODES) As Byte
Private Declare Function FSOUND_Close Lib "fmod.dll" Alias "_FSOUND_Close@0" () As Long
Private Declare Function FSOUND_Update Lib "fmod.dll" Alias "_FSOUND_Update@0" () As Long
Private Declare Function FSOUND_SetSpeakerMode Lib "fmod.dll" Alias "_FSOUND_SetSpeakerMode@4" (ByVal speakermode As FSOUND_SPEAKERMODES) As Long
Private Declare Function FSOUND_SetSFXMasterVolume Lib "fmod.dll" Alias "_FSOUND_SetSFXMasterVolume@4" (ByVal volume As Long) As Long
Private Declare Function FSOUND_SetPanSeperation Lib "fmod.dll" Alias "_FSOUND_SetPanSeperation@4" (ByVal pansep As Single) As Long
Private Declare Function FSOUND_File_SetCallbacks Lib "fmod.dll" Alias "_FSOUND_File_SetCallbacks@20" (ByVal OpenCallback As Long, ByVal CloseCallback As Long, ByVal ReadCallback As Long, ByVal SeekCallback As Long, ByVal TellCallback As Long) As Long
Private Declare Function FSOUND_GetError Lib "fmod.dll" Alias "_FSOUND_GetError@0" () As FMOD_ERRORS
Private Declare Function FSOUND_GetVersion Lib "fmod.dll" Alias "_FSOUND_GetVersion@0" () As Single
Private Declare Function FSOUND_GetOutput Lib "fmod.dll" Alias "_FSOUND_GetOutput@0" () As FSOUND_OUTPUTTYPES
Private Declare Function FSOUND_GetOutputHandle Lib "fmod.dll" Alias "_FSOUND_GetOutputHandle@0" () As Long
Private Declare Function FSOUND_GetDriver Lib "fmod.dll" Alias "_FSOUND_GetDriver@0" () As Long
Private Declare Function FSOUND_GetMixer Lib "fmod.dll" Alias "_FSOUND_GetMixer@0" () As FSOUND_MIXERTYPES
Private Declare Function FSOUND_GetNumDrivers Lib "fmod.dll" Alias "_FSOUND_GetNumDrivers@0" () As Long
Private Declare Function FSOUND_GetDriverName Lib "fmod.dll" Alias "_FSOUND_GetDriverName@4" (ByVal id As Long) As Long
Private Declare Function FSOUND_GetDriverCaps Lib "fmod.dll" Alias "_FSOUND_GetDriverCaps@8" (ByVal id As Long, ByRef caps As Long) As Byte
Private Declare Function FSOUND_GetOutputRate Lib "fmod.dll" Alias "_FSOUND_GetOutputRate@0" () As Long
Private Declare Function FSOUND_GetMaxChannels Lib "fmod.dll" Alias "_FSOUND_GetMaxChannels@0" () As Long
Private Declare Function FSOUND_GetMaxSamples Lib "fmod.dll" Alias "_FSOUND_GetMaxSamples@0" () As Long
Private Declare Function FSOUND_GetSFXMasterVolume Lib "fmod.dll" Alias "_FSOUND_GetSFXMasterVolume@0" () As Long
Private Declare Function FSOUND_GetNumHWChannels Lib "fmod.dll" Alias "_FSOUND_GetNumHWChannels@12" (ByRef num2d As Long, ByRef num3d As Long, ByRef total As Long)
Private Declare Function FSOUND_GetChannelsPlaying Lib "fmod.dll" Alias "_FSOUND_GetChannelsPlaying@0" () As Long
Private Declare Function FSOUND_GetCPUUsage Lib "fmod.dll" Alias "_FSOUND_GetCPUUsage@0" () As Single
Private Declare Sub FSOUND_GetMemoryStats Lib "fmod.dll" Alias "_FSOUND_GetMemoryStats@8" (ByRef currentalloced As Long, ByRef maxalloced As Long)
Private Declare Function FSOUND_Sample_Load Lib "fmod.dll" Alias "_FSOUND_Sample_Load@20" (ByVal index As Long, ByVal name As String, ByVal mode As FSOUND_MODES, ByVal offset As Long, ByVal length As Long) As Long
Private Declare Function FSOUND_Sample_Alloc Lib "fmod.dll" Alias "_FSOUND_Sample_Alloc@28" (ByVal index As Long, ByVal length As Long, ByVal mode As Long, ByVal deffreq As Long, ByVal defvol As Long, ByVal defpan As Long, ByVal defpri As Long) As Long
Private Declare Function FSOUND_Sample_Free Lib "fmod.dll" Alias "_FSOUND_Sample_Free@4" (ByVal sptr As Long) As Long
Private Declare Function FSOUND_Sample_Upload Lib "fmod.dll" Alias "_FSOUND_Sample_Upload@12" (ByVal sptr As Long, ByRef srcdata As Long, ByVal mode As Long) As Byte
Private Declare Function FSOUND_Sample_Lock Lib "fmod.dll" Alias "_FSOUND_Sample_Lock@28" (ByVal sptr As Long, ByVal offset As Long, ByVal length As Long, ByRef ptr1 As Long, ByRef ptr2 As Long, ByRef len1 As Long, ByRef len2 As Long) As Byte
Private Declare Function FSOUND_Sample_Unlock Lib "fmod.dll" Alias "_FSOUND_Sample_Unlock@20" (ByVal sptr As Long, ByVal sptr1 As Long, ByVal sptr2 As Long, ByVal len1 As Long, ByVal len2 As Long) As Byte
Private Declare Function FSOUND_Sample_SetMode Lib "fmod.dll" Alias "_FSOUND_Sample_SetMode@8" (ByVal sptr As Long, ByVal mode As FSOUND_MODES) As Byte
Private Declare Function FSOUND_Sample_SetLoopPoints Lib "fmod.dll" Alias "_FSOUND_Sample_SetLoopPoints@12" (ByVal sptr As Long, ByVal loopstart As Long, ByVal loopend As Long) As Byte
Private Declare Function FSOUND_Sample_SetDefaults Lib "fmod.dll" Alias "_FSOUND_Sample_SetDefaults@20" (ByVal sptr As Long, ByVal deffreq As Long, ByVal defvol As Long, ByVal defpan As Long, ByVal defpri As Long) As Byte
Private Declare Function FSOUND_Sample_SetDefaultsEx Lib "fmod.dll" Alias "_FSOUND_Sample_SetDefaultsEx@32" (ByVal sptr As Long, ByVal deffreq As Long, ByVal defvol As Long, ByVal defpan As Long, ByVal defpri As Long, ByVal varfreq As Long, ByVal varvol As Long, ByVal varpan As Long) As Byte
Private Declare Function FSOUND_Sample_SetMinMaxDistance Lib "fmod.dll" Alias "_FSOUND_Sample_SetMinMaxDistance@12" (ByVal sptr As Long, ByVal min As Single, ByVal max As Single) As Byte
Private Declare Function FSOUND_Sample_SetMaxPlaybacks Lib "fmod.dll" Alias "_FSOUND_Sample_SetMaxPlaybacks@8" (ByVal sptr As Long, ByVal max As Long) As Byte
Private Declare Function FSOUND_Sample_Get Lib "fmod.dll" Alias "_FSOUND_Sample_Get@4" (ByVal sampno As Long) As Long
Private Declare Function FSOUND_Sample_GetName Lib "fmod.dll" Alias "_FSOUND_Sample_GetName@4" (ByVal sptr As Long) As Long
Private Declare Function FSOUND_Sample_GetLength Lib "fmod.dll" Alias "_FSOUND_Sample_GetLength@4" (ByVal sptr As Long) As Long
Private Declare Function FSOUND_Sample_GetLoopPoints Lib "fmod.dll" Alias "_FSOUND_Sample_GetLoopPoints@12" (ByVal sptr As Long, ByRef loopstart As Long, ByRef loopend As Long) As Byte
Private Declare Function FSOUND_Sample_GetDefaults Lib "fmod.dll" Alias "_FSOUND_Sample_GetDefaults@20" (ByVal sptr As Long, ByRef deffreq As Long, ByRef defvol As Long, ByRef defpan As Long, ByRef defpri As Long) As Byte
Private Declare Function FSOUND_Sample_GetDefaultsEx Lib "fmod.dll" Alias "_FSOUND_Sample_GetDefaultsEx@32" (ByVal sptr As Long, ByRef deffreq As Long, ByRef defvol As Long, ByRef defpan As Long, ByRef defpri As Long, ByRef varfreq As Long, ByRef varvol As Long, ByRef varpan As Long) As Byte
Private Declare Function FSOUND_Sample_GetMode Lib "fmod.dll" Alias "_FSOUND_Sample_GetMode@4" (ByVal sptr As Long) As Long
Private Declare Function FSOUND_Sample_GetMinMaxDistance Lib "fmod.dll" Alias "_FSOUND_Sample_GetMinMaxDistance@12" (ByVal sptr As Long, ByRef min As Single, ByRef max As Single) As Byte
Private Declare Function FSOUND_PlaySound Lib "fmod.dll" Alias "_FSOUND_PlaySound@8" (ByVal channel As Long, ByVal sptr As Long) As Long
Private Declare Function FSOUND_PlaySoundEx Lib "fmod.dll" Alias "_FSOUND_PlaySoundEx@16" (ByVal channel As Long, ByVal sptr As Long, ByVal dsp As Long, ByVal startpaused As Byte) As Long
Private Declare Function FSOUND_StopSound Lib "fmod.dll" Alias "_FSOUND_StopSound@4" (ByVal channel As Long) As Byte
Private Declare Function FSOUND_SetFrequency Lib "fmod.dll" Alias "_FSOUND_SetFrequency@8" (ByVal channel As Long, ByVal freq As Long) As Byte
Private Declare Function FSOUND_SetVolume Lib "fmod.dll" Alias "_FSOUND_SetVolume@8" (ByVal channel As Long, ByVal Vol As Long) As Byte
Private Declare Function FSOUND_SetVolumeAbsolute Lib "fmod.dll" Alias "_FSOUND_SetVolumeAbsolute@8" (ByVal channel As Long, ByVal Vol As Long) As Byte
Private Declare Function FSOUND_SetPan Lib "fmod.dll" Alias "_FSOUND_SetPan@8" (ByVal channel As Long, ByVal pan As Long) As Byte
Private Declare Function FSOUND_SetSurround Lib "fmod.dll" Alias "_FSOUND_SetSurround@8" (ByVal channel As Long, ByVal surround As Long) As Byte
Private Declare Function FSOUND_SetMute Lib "fmod.dll" Alias "_FSOUND_SetMute@8" (ByVal channel As Long, ByVal mute As Byte) As Byte
Private Declare Function FSOUND_SetPriority Lib "fmod.dll" Alias "_FSOUND_SetPriority@8" (ByVal channel As Long, ByVal Priority As Long) As Byte
Private Declare Function FSOUND_SetReserved Lib "fmod.dll" Alias "_FSOUND_SetReserved@8" (ByVal channel As Long, ByVal reserved As Long) As Byte
Private Declare Function FSOUND_SetPaused Lib "fmod.dll" Alias "_FSOUND_SetPaused@8" (ByVal channel As Long, ByVal Paused As Byte) As Byte
Private Declare Function FSOUND_SetLoopMode Lib "fmod.dll" Alias "_FSOUND_SetLoopMode@8" (ByVal channel As Long, ByVal loopmode As Long) As Byte
Private Declare Function FSOUND_SetCurrentPosition Lib "fmod.dll" Alias "_FSOUND_SetCurrentPosition@8" (ByVal channel As Long, ByVal offset As Long) As Byte
Private Declare Function FSOUND_3D_SetAttributes Lib "fmod.dll" Alias "_FSOUND_3D_SetAttributes@12" (ByVal channel As Long, ByRef Pos As Single, ByRef vel As Single) As Byte
Private Declare Function FSOUND_3D_SetMinMaxDistance Lib "fmod.dll" Alias "_FSOUND_3D_SetMinMaxDistance@12" (ByVal channel As Long, ByVal min As Single, ByVal max As Single) As Byte
Private Declare Function FSOUND_IsPlaying Lib "fmod.dll" Alias "_FSOUND_IsPlaying@4" (ByVal channel As Long) As Byte
Private Declare Function FSOUND_GetFrequency Lib "fmod.dll" Alias "_FSOUND_GetFrequency@4" (ByVal channel As Long) As Long
Private Declare Function FSOUND_GetVolume Lib "fmod.dll" Alias "_FSOUND_GetVolume@4" (ByVal channel As Long) As Long
Private Declare Function FSOUND_GetAmplitude Lib "fmod.dll" Alias "_FSOUND_GetAmplitude@4" (ByVal channel As Long) As Long
Private Declare Function FSOUND_GetPan Lib "fmod.dll" Alias "_FSOUND_GetPan@4" (ByVal channel As Long) As Long
Private Declare Function FSOUND_GetSurround Lib "fmod.dll" Alias "_FSOUND_GetSurround@4" (ByVal channel As Long) As Byte
Private Declare Function FSOUND_GetMute Lib "fmod.dll" Alias "_FSOUND_GetMute@4" (ByVal channel As Long) As Byte
Private Declare Function FSOUND_GetPriority Lib "fmod.dll" Alias "_FSOUND_GetPriority@4" (ByVal channel As Long) As Long
Private Declare Function FSOUND_GetReserved Lib "fmod.dll" Alias "_FSOUND_GetReserved@4" (ByVal channel As Long) As Byte
Private Declare Function FSOUND_GetPaused Lib "fmod.dll" Alias "_FSOUND_GetPaused@4" (ByVal channel As Long) As Byte
Private Declare Function FSOUND_GetLoopMode Lib "fmod.dll" Alias "_FSOUND_GetLoopMode@4" (ByVal channel As Long) As Long
Private Declare Function FSOUND_GetCurrentPosition Lib "fmod.dll" Alias "_FSOUND_GetCurrentPosition@4" (ByVal channel As Long) As Long
Private Declare Function FSOUND_GetCurrentSample Lib "fmod.dll" Alias "_FSOUND_GetCurrentSample@4" (ByVal channel As Long) As Long
Private Declare Function FSOUND_GetCurrentLevels Lib "fmod.dll" Alias "_FSOUND_GetCurrentLevels@12" (ByVal channel As Long, ByRef l As Single, ByRef r As Single) As Byte
Private Declare Function FSOUND_GetNumSubChannels Lib "fmod.dll" Alias "_FSOUND_GetNumSubChannels@4" (ByVal channel As Long) As Long
Private Declare Function FSOUND_GetSubChannel Lib "fmod.dll" Alias "_FSOUND_GetSubChannel@8" (ByVal channel As Long, ByVal subchannel As Long) As Long
Private Declare Function FSOUND_3D_GetAttributes Lib "fmod.dll" Alias "_FSOUND_3D_GetAttributes@12" (ByVal channel As Long, ByRef Pos As Single, ByRef vel As Single) As Byte
Private Declare Function FSOUND_3D_GetMinMaxDistance Lib "fmod.dll" Alias "_FSOUND_3D_GetMinMaxDistance@12" (ByVal channel As Long, ByRef min As Single, ByRef max As Single) As Byte
Private Declare Function FSOUND_3D_Listener_SetCurrent Lib "fmod.dll" Alias "_FSOUND_3D_Listener_SetCurrent@8" (ByVal current As Long) As Long
Private Declare Function FSOUND_3D_Listener_SetAttributes Lib "fmod.dll" Alias "_FSOUND_3D_Listener_SetAttributes@32" (ByVal Pos As Single, ByVal vel As Single, ByVal fx As Single, ByVal fy As Single, ByVal fz As Single, ByVal tx As Single, ByVal ty As Single, ByVal tz As Single) As Long
Private Declare Function FSOUND_3D_Listener_GetAttributes Lib "fmod.dll" Alias "_FSOUND_3D_Listener_GetAttributes@32" (ByRef Pos As Single, ByRef vel As Single, ByRef fx As Single, ByRef fy As Single, ByRef fz As Single, ByRef tx As Single, ByRef ty As Single, ByRef tz As Single) As Long
Private Declare Function FSOUND_3D_SetDopplerFactor Lib "fmod.dll" Alias "_FSOUND_3D_SetDopplerFactor@4" (ByVal fscale As Single) As Long
Private Declare Function FSOUND_3D_SetDistanceFactor Lib "fmod.dll" Alias "_FSOUND_3D_SetDistanceFactor@4" (ByVal fscale As Single) As Long
Private Declare Function FSOUND_3D_SetRolloffFactor Lib "fmod.dll" Alias "_FSOUND_3D_SetRolloffFactor@4" (ByVal fscale As Single) As Long
Private Declare Function FSOUND_FX_Enable Lib "fmod.dll" Alias "_FSOUND_FX_Enable@8" (ByVal channel As Long, ByVal fx As FSOUND_FX_MODES) As Long
Private Declare Function FSOUND_FX_Disable Lib "fmod.dll" Alias "_FSOUND_FX_Disable@4" (ByVal channel As Long) As Byte
Private Declare Function FSOUND_FX_SetChorus Lib "fmod.dll" Alias "_FSOUND_FX_SetChorus@32" (ByVal fxid As Long, ByVal WetDryMix As Single, ByVal Depth As Single, ByVal Feedback As Single, ByVal Frequency As Single, ByVal Waveform As Long, ByVal Delay As Single, ByVal Phase As Long) As Byte
Private Declare Function FSOUND_FX_SetCompressor Lib "fmod.dll" Alias "_FSOUND_FX_SetCompressor@28" (ByVal fxid As Long, ByVal Gain As Single, ByVal Attack As Single, ByVal Release As Single, ByVal Threshold As Single, ByVal Ratio As Single, ByVal Predelay As Single) As Byte
Private Declare Function FSOUND_FX_SetDistortion Lib "fmod.dll" Alias "_FSOUND_FX_SetDistortion@24" (ByVal fxid As Long, ByVal Gain As Single, ByVal Edge As Single, ByVal PostEQCenterFrequency As Single, ByVal PostEQBandwidth As Single, ByVal PreLowpassCutoff As Single) As Byte
Private Declare Function FSOUND_FX_SetEcho Lib "fmod.dll" Alias "_FSOUND_FX_SetEcho@24" (ByVal fxid As Long, ByVal WetDryMix As Single, ByVal Feedback As Single, ByVal LeftDelay As Single, ByVal RightDelay As Single, ByVal PanDelay As Long) As Byte
Private Declare Function FSOUND_FX_SetFlanger Lib "fmod.dll" Alias "_FSOUND_FX_SetFlanger@32" (ByVal fxid As Long, ByVal WetDryMix As Single, ByVal Depth As Single, ByVal Feedback As Single, ByVal Frequency As Single, ByVal Waveform As Long, ByVal Delay As Single, ByVal Phase As Long) As Byte
Private Declare Function FSOUND_FX_SetGargle Lib "fmod.dll" Alias "_FSOUND_FX_SetGargle@12" (ByVal fxid As Long, ByVal RateHz As Long, ByVal WaveShape As Long) As Byte
Private Declare Function FSOUND_FX_SetI3DL2Reverb Lib "fmod.dll" Alias "_FSOUND_FX_SetI3DL2Reverb@52" (ByVal fxid As Long, ByVal Room As Long, ByVal RoomHF As Long, ByVal RoomRolloffFactor As Single, ByVal DecayTime As Single, ByVal DecayHFRatio As Single, ByVal Reflections As Long, ByVal ReflectionsDelay As Single, ByVal Reverb As Long, ByVal ReverbDelay As Single, ByVal Diffusion As Single, ByVal Density As Single, ByVal HFReference As Single) As Byte
Private Declare Function FSOUND_FX_SetParamEQ Lib "fmod.dll" Alias "_FSOUND_FX_SetParamEQ@16" (ByVal fxid As Long, ByVal Center As Single, ByVal Bandwidth As Single, ByVal Gain As Single) As Byte
Private Declare Function FSOUND_FX_SetWavesReverb Lib "fmod.dll" Alias "_FSOUND_FX_SetWavesReverb@20" (ByVal fxid As Long, ByVal InGain As Single, ByVal ReverbMix As Single, ByVal ReverbTime As Single, ByVal HighFreqRTRatio As Single) As Byte
Private Declare Function FSOUND_Stream_SetBufferSize Lib "fmod.dll" Alias "_FSOUND_Stream_SetBufferSize@4" (ByVal ms As Long) As Byte
Private Declare Function FSOUND_Stream_Open Lib "fmod.dll" Alias "_FSOUND_Stream_Open@16" (ByVal filename As String, ByVal mode As FSOUND_MODES, ByVal offset As Long, ByVal length As Long) As Long
Private Declare Function FSOUND_Stream_Open2 Lib "fmod.dll" Alias "_FSOUND_Stream_Open@16" (ByRef data As Byte, ByVal mode As FSOUND_MODES, ByVal offset As Long, ByVal length As Long) As Long
Private Declare Function FSOUND_Stream_Create Lib "fmod.dll" Alias "_FSOUND_Stream_Create@20" (ByVal callback As Long, ByVal length As Long, ByVal mode As Long, ByVal samplerate As Long, ByVal userdata As Long) As Long
Private Declare Function FSOUND_Stream_Close Lib "fmod.dll" Alias "_FSOUND_Stream_Close@4" (ByVal stream As Long) As Byte
Private Declare Function FSOUND_Stream_Play Lib "fmod.dll" Alias "_FSOUND_Stream_Play@8" (ByVal channel As Long, ByVal stream As Long) As Long
Private Declare Function FSOUND_Stream_PlayEx Lib "fmod.dll" Alias "_FSOUND_Stream_PlayEx@16" (ByVal channel As Long, ByVal stream As Long, ByVal dsp As Long, ByVal startpaused As Byte) As Long
Private Declare Function FSOUND_Stream_Stop Lib "fmod.dll" Alias "_FSOUND_Stream_Stop@4" (ByVal stream As Long) As Byte
Private Declare Function FSOUND_Stream_SetPosition Lib "fmod.dll" Alias "_FSOUND_Stream_SetPosition@8" (ByVal stream As Long, ByVal positition As Long) As Byte
Private Declare Function FSOUND_Stream_GetPosition Lib "fmod.dll" Alias "_FSOUND_Stream_GetPosition@4" (ByVal stream As Long) As Long
Private Declare Function FSOUND_Stream_SetTime Lib "fmod.dll" Alias "_FSOUND_Stream_SetTime@8" (ByVal stream As Long, ByVal ms As Long) As Byte
Private Declare Function FSOUND_Stream_GetTime Lib "fmod.dll" Alias "_FSOUND_Stream_GetTime@4" (ByVal stream As Long) As Long
Private Declare Function FSOUND_Stream_GetLength Lib "fmod.dll" Alias "_FSOUND_Stream_GetLength@4" (ByVal stream As Long) As Long
Private Declare Function FSOUND_Stream_GetLengthMs Lib "fmod.dll" Alias "_FSOUND_Stream_GetLengthMs@4" (ByVal stream As Long) As Long
Private Declare Function FSOUND_Stream_SetMode Lib "fmod.dll" Alias "_FSOUND_Stream_SetMode@8" (ByVal stream As Long, ByVal mode As Long) As Byte
Private Declare Function FSOUND_Stream_GetMode Lib "fmod.dll" Alias "_FSOUND_Stream_GetMode@4" (ByVal stream As Long) As Long
Private Declare Function FSOUND_Stream_SetLoopPoints Lib "fmod.dll" Alias "_FSOUND_Stream_SetLoopPoints@12" (ByVal stream As Long, ByVal loopstartpcm As Long, ByVal loopendpcm As Long) As Byte
Private Declare Function FSOUND_Stream_SetLoopCount Lib "fmod.dll" Alias "_FSOUND_Stream_SetLoopCount@8" (ByVal stream As Long, ByVal count As Long) As Byte
Private Declare Function FSOUND_Stream_GetOpenState Lib "fmod.dll" Alias "_FSOUND_Stream_GetOpenState@4" (ByVal stream As Long) As Long
Private Declare Function FSOUND_Stream_GetSample Lib "fmod.dll" Alias "_FSOUND_Stream_GetSample@4" (ByVal stream As Long) As Long
Private Declare Function FSOUND_Stream_CreateDSP Lib "fmod.dll" Alias "_FSOUND_Stream_CreateDSP@16" (ByVal stream As Long, ByVal callback As Long, ByVal Priority As Long, ByVal userdata As Long) As Long
Private Declare Function FSOUND_Stream_SetEndCallback Lib "fmod.dll" Alias "_FSOUND_Stream_SetEndCallback@12" (ByVal stream As Long, ByVal callback As Long, ByVal userdata As Long) As Byte
Private Declare Function FSOUND_Stream_SetSyncCallback Lib "fmod.dll" Alias "_FSOUND_Stream_SetSyncCallback@12" (ByVal stream As Long, ByVal callback As Long, ByVal userdata As Long) As Byte
Private Declare Function FSOUND_Stream_AddSyncPoint Lib "fmod.dll" Alias "_FSOUND_Stream_AddSyncPoint@12" (ByVal stream As Long, ByVal pcmoffset As Long, ByVal name As String) As Long
Private Declare Function FSOUND_Stream_DeleteSyncPoint Lib "fmod.dll" Alias "_FSOUND_Stream_DeleteSyncPoint@4" (ByVal point As Long) As Byte
Private Declare Function FSOUND_Stream_GetNumSyncPoints Lib "fmod.dll" Alias "_FSOUND_Stream_GetNumSyncPoints@4" (ByVal stream As Long) As Long
Private Declare Function FSOUND_Stream_GetSyncPoint Lib "fmod.dll" Alias "_FSOUND_Stream_GetSyncPoint@8" (ByVal stream As Long, ByVal index As Long) As Long
Private Declare Function FSOUND_Stream_GetSyncPointInfo Lib "fmod.dll" Alias "_FSOUND_Stream_GetSyncPointInfo@8" (ByVal point As Long, ByRef pcmoffset As Long) As Long
Private Declare Function FSOUND_Stream_SetSubStream Lib "fmod.dll" Alias "_FSOUND_Stream_SetSubStream@8" (ByVal stream As Long, ByVal index As Long) As Byte
Private Declare Function FSOUND_Stream_GetNumSubStreams Lib "fmod.dll" Alias "_FSOUND_Stream_GetNumSubStreams@4" (ByVal stream As Long) As Long
Private Declare Function FSOUND_Stream_SetSubStreamSentence Lib "fmod.dll" Alias "_FSOUND_Stream_SetSubStreamSentence@12" (ByVal stream As Long, ByRef sentencelist As Long, ByVal numitems As Long) As Byte
Private Declare Function FSOUND_Stream_GetNumTagFields Lib "fmod.dll" Alias "_FSOUND_Stream_GetNumTagFields@8" (ByVal stream As Long, ByRef num As Long) As Byte
Private Declare Function FSOUND_Stream_GetTagField Lib "fmod.dll" Alias "_FSOUND_Stream_GetTagField@24" (ByVal stream As Long, ByVal num As Long, ByRef tagtype As Long, ByRef name As Long, ByRef value As Long, ByRef length As Long) As Byte
Private Declare Function FSOUND_Stream_FindTagField Lib "fmod.dll" Alias "_FSOUND_Stream_FindTagField@20" (ByVal stream As Long, ByVal tagtype As Long, ByVal name As String, ByRef value As Long, ByRef length As Long) As Byte
Private Declare Function FSOUND_Stream_Net_SetProxy Lib "fmod.dll" Alias "_FSOUND_Stream_Net_SetProxy@4" (ByVal proxy As String) As Byte
Private Declare Function FSOUND_Stream_Net_GetLastServerStatus Lib "fmod.dll" Alias "_FSOUND_Stream_Net_GetLastServerStatus@0" () As Long
Private Declare Function FSOUND_Stream_Net_SetBufferProperties Lib "fmod.dll" Alias "_FSOUND_Stream_Net_SetBufferProperties@12" (ByVal buffersize As Long, ByVal prebuffer_percent As Long, ByVal rebuffer_percent As Long) As Byte
Private Declare Function FSOUND_Stream_Net_GetBufferProperties Lib "fmod.dll" Alias "_FSOUND_Stream_Net_GetBufferProperties@12" (ByRef buffersize As Long, ByRef prebuffer_percent As Long, ByRef rebuffer_percent As Long) As Byte
Private Declare Function FSOUND_Stream_Net_SetMetadataCallback Lib "fmod.dll" Alias "_FSOUND_Stream_Net_SetMetadataCallback@12" (ByVal stream As Long, ByVal callback As Long, ByVal userdata As Long) As Byte
Private Declare Function FSOUND_Stream_Net_GetStatus Lib "fmod.dll" Alias "_FSOUND_Stream_Net_GetStatus@20" (ByVal stream As Long, ByRef status As Long, ByRef bufferpercentused As Long, ByRef bitrate As Long, ByRef flags As Long) As Byte
Private Declare Function FSOUND_CD_Play Lib "fmod.dll" Alias "_FSOUND_CD_Play@8" (ByVal drive As Byte, ByVal Track As Long) As Byte
Private Declare Function FSOUND_CD_SetPlayMode Lib "fmod.dll" Alias "_FSOUND_CD_SetPlayMode@8" (ByVal drive As Byte, ByVal mode As FSOUND_CDPLAYMODES) As Long
Private Declare Function FSOUND_CD_Stop Lib "fmod.dll" Alias "_FSOUND_CD_Stop@4" (ByVal drive As Byte) As Byte
Private Declare Function FSOUND_CD_SetPaused Lib "fmod.dll" Alias "_FSOUND_CD_SetPaused@8" (ByVal drive As Byte, ByVal Paused As Byte) As Byte
Private Declare Function FSOUND_CD_SetVolume Lib "fmod.dll" Alias "_FSOUND_CD_SetVolume@8" (ByVal drive As Byte, ByVal volume As Long) As Byte
Private Declare Function FSOUND_CD_SetTrackTime Lib "fmod.dll" Alias "_FSOUND_CD_SetTrackTime@8" (ByVal drive As Byte, ByVal ms As Long) As Byte
Private Declare Function FSOUND_CD_OpenTray Lib "fmod.dll" Alias "_FSOUND_CD_OpenTray@8" (ByVal drive As Byte, ByVal openState As Byte) As Byte
Private Declare Function FSOUND_CD_GetPaused Lib "fmod.dll" Alias "_FSOUND_CD_GetPaused@4" (ByVal drive As Byte) As Byte
Private Declare Function FSOUND_CD_GetTrack Lib "fmod.dll" Alias "_FSOUND_CD_GetTrack@4" (ByVal drive As Byte) As Long
Private Declare Function FSOUND_CD_GetNumTracks Lib "fmod.dll" Alias "_FSOUND_CD_GetNumTracks@4" (ByVal drive As Byte) As Long
Private Declare Function FSOUND_CD_GetVolume Lib "fmod.dll" Alias "_FSOUND_CD_GetVolume@4" (ByVal drive As Byte) As Long
Private Declare Function FSOUND_CD_GetTrackLength Lib "fmod.dll" Alias "_FSOUND_CD_GetTrackLength@8" (ByVal drive As Byte, ByVal Track As Long) As Long
Private Declare Function FSOUND_CD_GetTrackTime Lib "fmod.dll" Alias "_FSOUND_CD_GetTrackTime@4" (ByVal drive As Byte) As Long
Private Declare Function FSOUND_DSP_Create Lib "fmod.dll" Alias "_FSOUND_DSP_Create@12" (ByVal callback As Long, ByVal Priority As Long, ByVal param As Long) As Long
Private Declare Function FSOUND_DSP_Free Lib "fmod.dll" Alias "_FSOUND_DSP_Free@4" (ByVal unit As Long) As Long
Private Declare Function FSOUND_DSP_SetPriority Lib "fmod.dll" Alias "_FSOUND_DSP_SetPriority@8" (ByVal unit As Long, ByVal Priority As Long) As Long
Private Declare Function FSOUND_DSP_GetPriority Lib "fmod.dll" Alias "_FSOUND_DSP_GetPriority@4" (ByVal unit As Long) As Long
Private Declare Function FSOUND_DSP_SetActive Lib "fmod.dll" Alias "_FSOUND_DSP_SetActive@8" (ByVal unit As Long, ByVal active As Integer) As Long
Private Declare Function FSOUND_DSP_GetActive Lib "fmod.dll" Alias "_FSOUND_DSP_GetActive@4" (ByVal unit As Long) As Byte
Private Declare Function FSOUND_DSP_GetClearUnit Lib "fmod.dll" Alias "_FSOUND_DSP_GetClearUnit@0" () As Long
Private Declare Function FSOUND_DSP_GetSFXUnit Lib "fmod.dll" Alias "_FSOUND_DSP_GetSFXUnit@0" () As Long
Private Declare Function FSOUND_DSP_GetMusicUnit Lib "fmod.dll" Alias "_FSOUND_DSP_GetMusicUnit@0" () As Long
Private Declare Function FSOUND_DSP_GetFFTUnit Lib "fmod.dll" Alias "_FSOUND_DSP_GetFFTUnit@0" () As Long
Private Declare Function FSOUND_DSP_GetClipAndCopyUnit Lib "fmod.dll" Alias "_FSOUND_DSP_GetClipAndCopyUnit@0" () As Long
Private Declare Function FSOUND_DSP_MixBuffers Lib "fmod.dll" Alias "_FSOUND_DSP_MixBuffers@28" (ByVal destbuffer As Long, ByVal srcbuffer As Long, ByVal length As Long, ByVal freq As Long, ByVal Vol As Long, ByVal pan As Long, ByVal mode As Long) As Byte
Private Declare Function FSOUND_DSP_ClearMixBuffer Lib "fmod.dll" Alias "_FSOUND_DSP_ClearMixBuffer@0" () As Long
Private Declare Function FSOUND_DSP_GetBufferLength Lib "fmod.dll" Alias "_FSOUND_DSP_GetBufferLength@0" () As Long
Private Declare Function FSOUND_DSP_GetBufferLengthTotal Lib "fmod.dll" Alias "_FSOUND_DSP_GetBufferLengthTotal@0" () As Long
Private Declare Function FSOUND_DSP_GetSpectrum Lib "fmod.dll" Alias "_FSOUND_DSP_GetSpectrum@0" () As Long
Private Declare Function FSOUND_Reverb_SetProperties Lib "fmod.dll" Alias "_FSOUND_Reverb_SetProperties@4" (ByRef prop As FSOUND_REVERB_PROPERTIES) As Byte
Private Declare Function FSOUND_Reverb_GetProperties Lib "fmod.dll" Alias "_FSOUND_Reverb_GetProperties@4" (ByRef prop As FSOUND_REVERB_PROPERTIES) As Byte
Private Declare Function FSOUND_Reverb_SetChannelProperties Lib "fmod.dll" Alias "_FSOUND_Reverb_SetChannelProperties@8" (ByVal channel As Long, ByRef prop As FSOUND_REVERB_CHANNELPROPERTIES) As Byte
Private Declare Function FSOUND_Reverb_GetChannelProperties Lib "fmod.dll" Alias "_FSOUND_Reverb_GetChannelProperties@8" (ByVal channel As Long, ByRef prop As FSOUND_REVERB_CHANNELPROPERTIES) As Byte
Private Declare Function FSOUND_Record_SetDriver Lib "fmod.dll" Alias "_FSOUND_Record_SetDriver@4" (ByVal outputtype As Long) As Byte
Private Declare Function FSOUND_Record_GetNumDrivers Lib "fmod.dll" Alias "_FSOUND_Record_GetNumDrivers@0" () As Long
Private Declare Function FSOUND_Record_GetDriverName Lib "fmod.dll" Alias "_FSOUND_Record_GetDriverName@4" (ByVal id As Long) As Long
Private Declare Function FSOUND_Record_GetDriver Lib "fmod.dll" Alias "_FSOUND_Record_GetDriver@0" () As Long
Private Declare Function FSOUND_Record_StartSample Lib "fmod.dll" Alias "_FSOUND_Record_StartSample@8" (ByVal sample As Long, ByVal loopit As Boolean) As Byte
Private Declare Function FSOUND_Record_Stop Lib "fmod.dll" Alias "_FSOUND_Record_Stop@0" () As Byte
Private Declare Function FSOUND_Record_GetPosition Lib "fmod.dll" Alias "_FSOUND_Record_GetPosition@0" () As Long
Private Declare Function FMUSIC_LoadSong Lib "fmod.dll" Alias "_FMUSIC_LoadSong@4" (ByVal name As String) As Long
Private Declare Function FMUSIC_LoadSongEx Lib "fmod.dll" Alias "_FMUSIC_LoadSongEx@24" (ByVal name As String, ByVal offset As Long, ByVal length As Long, ByVal mode As FSOUND_MODES, ByRef sentencelist As Long, ByVal numitems As Long) As Long
Private Declare Function FMUSIC_LoadSongEx2 Lib "fmod.dll" Alias "_FMUSIC_LoadSongEx@24" (ByRef data As Byte, ByVal offset As Long, ByVal length As Long, ByVal mode As FSOUND_MODES, ByRef sentencelist As Long, ByVal numitems As Long) As Long
Private Declare Function FMUSIC_GetOpenState Lib "fmod.dll" Alias "_FMUSIC_GetOpenState@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_FreeSong Lib "fmod.dll" Alias "_FMUSIC_FreeSong@4" (ByVal module As Long) As Byte
Private Declare Function FMUSIC_PlaySong Lib "fmod.dll" Alias "_FMUSIC_PlaySong@4" (ByVal module As Long) As Byte
Private Declare Function FMUSIC_StopSong Lib "fmod.dll" Alias "_FMUSIC_StopSong@4" (ByVal module As Long) As Byte
Private Declare Function FMUSIC_StopAllSongs Lib "fmod.dll" Alias "_FMUSIC_StopAllSongs@0" () As Long
Private Declare Function FMUSIC_SetZxxCallback Lib "fmod.dll" Alias "_FMUSIC_SetZxxCallback@8" (ByVal module As Long, ByVal callback As Long) As Byte
Private Declare Function FMUSIC_SetRowCallback Lib "fmod.dll" Alias "_FMUSIC_SetRowCallback@12" (ByVal module As Long, ByVal callback As Long, ByVal rowstep As Long) As Byte
Private Declare Function FMUSIC_SetOrderCallback Lib "fmod.dll" Alias "_FMUSIC_SetOrderCallback@12" (ByVal module As Long, ByVal callback As Long, ByVal rowstep As Long) As Byte
Private Declare Function FMUSIC_SetInstCallback Lib "fmod.dll" Alias "_FMUSIC_SetInstCallback@12" (ByVal module As Long, ByVal callback As Long, ByVal instrument As Long) As Byte
Private Declare Function FMUSIC_SetSample Lib "fmod.dll" Alias "_FMUSIC_SetSample@12" (ByVal module As Long, ByVal sampno As Long, ByVal sptr As Long) As Byte
Private Declare Function FMUSIC_SetUserData Lib "fmod.dll" Alias "_FMUSIC_SetUserData@8" (ByVal module As Long, ByVal userdata As Long) As Byte
Private Declare Function FMUSIC_OptimizeChannels Lib "fmod.dll" Alias "_FMUSIC_OptimizeChannels@12" (ByVal module As Long, ByVal maxchannels As Long, ByVal minvolume As Long) As Byte
Private Declare Function FMUSIC_SetReverb Lib "fmod.dll" Alias "_FMUSIC_SetReverb@4" (ByVal Reverb As Byte) As Byte
Private Declare Function FMUSIC_SetLooping Lib "fmod.dll" Alias "_FMUSIC_SetLooping@8" (ByVal module As Long, ByVal looping As Byte) As Byte
Private Declare Function FMUSIC_SetOrder Lib "fmod.dll" Alias "_FMUSIC_SetOrder@8" (ByVal module As Long, ByVal order As Long) As Byte
Private Declare Function FMUSIC_SetPaused Lib "fmod.dll" Alias "_FMUSIC_SetPaused@8" (ByVal module As Long, ByVal Pause As Byte) As Byte
Private Declare Function FMUSIC_SetMasterVolume Lib "fmod.dll" Alias "_FMUSIC_SetMasterVolume@8" (ByVal module As Long, ByVal volume As Long) As Byte
Private Declare Function FMUSIC_SetMasterSpeed Lib "fmod.dll" Alias "_FMUSIC_SetMasterSpeed@8" (ByVal module As Long, ByVal speed As Single) As Byte
Private Declare Function FMUSIC_SetPanSeperation Lib "fmod.dll" Alias "_FMUSIC_SetPanSeperation@8" (ByVal module As Long, ByVal pansep As Single) As Byte
Private Declare Function FMUSIC_GetName Lib "fmod.dll" Alias "_FMUSIC_GetName@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetType Lib "fmod.dll" Alias "_FMUSIC_GetType@4" (ByVal module As Long) As FMUSIC_TYPES
Private Declare Function FMUSIC_GetNumOrders Lib "fmod.dll" Alias "_FMUSIC_GetNumOrders@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetNumPatterns Lib "fmod.dll" Alias "_FMUSIC_GetNumPatterns@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetNumInstruments Lib "fmod.dll" Alias "_FMUSIC_GetNumInstruments@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetNumSamples Lib "fmod.dll" Alias "_FMUSIC_GetNumSamples@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetNumChannels Lib "fmod.dll" Alias "_FMUSIC_GetNumChannels@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetSample Lib "fmod.dll" Alias "_FMUSIC_GetSample@8" (ByVal module As Long, ByVal sampno As Long) As Long
Private Declare Function FMUSIC_GetPatternLength Lib "fmod.dll" Alias "_FMUSIC_GetPatternLength@8" (ByVal module As Long, ByVal orderno As Long) As Long
Private Declare Function FMUSIC_IsFinished Lib "fmod.dll" Alias "_FMUSIC_IsFinished@4" (ByVal module As Long) As Byte
Private Declare Function FMUSIC_IsPlaying Lib "fmod.dll" Alias "_FMUSIC_IsPlaying@4" (ByVal module As Long) As Byte
Private Declare Function FMUSIC_GetMasterVolume Lib "fmod.dll" Alias "_FMUSIC_GetMasterVolume@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetGlobalVolume Lib "fmod.dll" Alias "_FMUSIC_GetGlobalVolume@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetOrder Lib "fmod.dll" Alias "_FMUSIC_GetOrder@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetPattern Lib "fmod.dll" Alias "_FMUSIC_GetPattern@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetSpeed Lib "fmod.dll" Alias "_FMUSIC_GetSpeed@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetBPM Lib "fmod.dll" Alias "_FMUSIC_GetBPM@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetRow Lib "fmod.dll" Alias "_FMUSIC_GetRow@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetPaused Lib "fmod.dll" Alias "_FMUSIC_GetPaused@4" (ByVal module As Long) As Byte
Private Declare Function FMUSIC_GetTime Lib "fmod.dll" Alias "_FMUSIC_GetTime@4" (ByVal module As Long) As Long
Private Declare Function FMUSIC_GetRealChannel Lib "fmod.dll" Alias "_FMUSIC_GetRealChannel@8" (ByVal module As Long, ByVal modchannel As Long) As Long
Private Declare Function FMUSIC_GetUserData Lib "fmod.dll" Alias "_FMUSIC_GetUserData@4" (ByVal module As Long) As Long
Private Declare Function ConvCStringToVBString Lib "kernel32" Alias "lstrcpyA" (ByVal lpsz As String, ByVal pt As Long) As Long ' Notice the As Long return value replacing the As String given by the API Viewer.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Public Function FS_GetCPUUseage() As Single
On Local Error GoTo ErrHandler
FS_GetCPUUseage = FSOUND_GetCPUUsage()
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_SetMixer(lType As FSOUND_MIXERTYPES) As Byte
On Local Error GoTo ErrHandler
FS_SetMixer = FSOUND_SetMixer(lType)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_SetVolume(lChannel As Long, lVolume As Long) As Byte
On Local Error GoTo ErrHandler
FS_SetVolume = FSOUND_SetVolume(lChannel, lVolume)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_SetFrequency(lChannel As Long, lFrequency As Long) As Byte
On Local Error GoTo ErrHandler
FS_SetFrequency = FSOUND_SetFrequency(lChannel, lFrequency)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_GetOutput() As FSOUND_OUTPUTTYPES
On Local Error GoTo ErrHandler
FS_GetOutput = FSOUND_GetOutput()
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_SetPanSeparation(lPanSep As Single) As Byte
On Local Error GoTo ErrHandler
FS_SetPanSeparation = FSOUND_SetPanSeperation(lPanSep)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_SetBufferSize(lLenMS As Long) As Byte
On Local Error GoTo ErrHandler
FS_SetBufferSize = FSOUND_SetBufferSize(lLenMS)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_SetDriver(lDriver As Long) As Byte
On Local Error GoTo ErrHandler
FS_SetDriver = FSOUND_SetDriver(lDriver)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_SetOutput(lOutput As FSOUND_OUTPUTTYPES) As Byte
On Local Error GoTo ErrHandler
FS_SetOutput = FSOUND_SetOutput(lOutput)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_SetPaused(lPaused As Byte) As Byte
On Local Error GoTo ErrHandler
FS_SetPaused = FSOUND_SetPaused(lStreamChannel, lPaused)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_Stream_SetPosition(lPosition As Long) As Long
On Local Error GoTo ErrHandler
FS_Stream_SetPosition = FSOUND_Stream_SetPosition(lStreamHandle, lPosition)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_Stream_GetPosition() As Long
On Local Error GoTo ErrHandler
FS_Stream_GetPosition = FSOUND_Stream_GetPosition(lStreamHandle)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FM_SetReverb(lReverb As Byte) As Byte
On Local Error GoTo ErrHandler
FM_SetReverb = FMUSIC_SetReverb(lReverb)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_SetSpeakerMode(lSpeakerMode As FSOUND_SPEAKERMODES) As Long
On Local Error GoTo ErrHandler
FS_SetSpeakerMode = FSOUND_SetSpeakerMode(lSpeakerMode)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_SetSFXMasterVol(lPercent As Long) As Long
On Local Error GoTo ErrHandler
FS_SetSFXMasterVol = FSOUND_SetSFXMasterVolume(lPercent)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Private Function FSOUND_GetErrorString(ByVal errorcode As Long) As String
On Local Error GoTo ErrHandler
Select Case errorcode
Case FMOD_ERR_NONE
    FSOUND_GetErrorString = "No errors present"
Case FMOD_ERR_BUSY
    FSOUND_GetErrorString = "Cannot use this command after initializing, you must use ctlMusicPlayer1.Close first."
Case FMOD_ERR_UNINITIALIZED
    FSOUND_GetErrorString = "This command failed because FSOUND_Init was not called"
Case FMOD_ERR_PLAY
    FSOUND_GetErrorString = "Playing the sound failed."
Case FMOD_ERR_INIT
    FSOUND_GetErrorString = "Error initializing output device."
Case FMOD_ERR_ALLOCATED
    FSOUND_GetErrorString = "The output device is already in use and cannot be reused."
Case FMOD_ERR_OUTPUT_FORMAT
    FSOUND_GetErrorString = "Soundcard does not support the features needed for this soundsystem (16bit stereo output)"
Case FMOD_ERR_COOPERATIVELEVEL
    FSOUND_GetErrorString = "Error setting cooperative level for hardware."
Case FMOD_ERR_CREATEBUFFER
    FSOUND_GetErrorString = "Error creating hardware sound buffer."
Case FMOD_ERR_FILE_NOTFOUND
    FSOUND_GetErrorString = "File not found"
Case FMOD_ERR_FILE_FORMAT
    FSOUND_GetErrorString = "Unknown file format"
Case FMOD_ERR_FILE_BAD
    FSOUND_GetErrorString = "Error loading file"
Case FMOD_ERR_MEMORY
    FSOUND_GetErrorString = "Not enough memory"
Case FMOD_ERR_VERSION
    FSOUND_GetErrorString = "The version number of this file format is not supported"
Case FMOD_ERR_INVALID_PARAM
    FSOUND_GetErrorString = "An invalid parameter was passed to this function"
Case FMOD_ERR_NO_EAX
    FSOUND_GetErrorString = "Tried to use an EAX command on a non EAX enabled channel or output."
Case FMOD_ERR_CHANNEL_ALLOC
    FSOUND_GetErrorString = "Failed to allocate a new channel"
Case FMOD_ERR_RECORD
    FSOUND_GetErrorString = "Recording is not supported on this machine"
Case FMOD_ERR_MEDIAPLAYER
    FSOUND_GetErrorString = "Required Mediaplayer codec is not installed"
Case FMOD_ERR_CDDEVICE
    FSOUND_GetErrorString = "An error occured trying to open the specified CD device"
Case Else
    FSOUND_GetErrorString = "Unknown error"
End Select
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_Stream_Play() As Boolean
On Local Error GoTo ErrHandler
lStreamChannel = FSOUND_Stream_Play(FSOUND_FREE, lStreamHandle)
If lStreamChannel <> 0 Then
    FS_Stream_Play = True
Else
    FS_Stream_Play = False
End If
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Sub FS_Stream_Close()
On Local Error GoTo ErrHandler
FSOUND_Stream_Close lStreamHandle
lStreamHandle = 0
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Public Function FS_Stream_Stop()
On Local Error GoTo ErrHandler
FSOUND_Stream_Stop lStreamHandle
lStreamChannel = 0
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_ReturnLastError() As String
On Local Error GoTo ErrHandler
FS_ReturnLastError = FSOUND_GetErrorString(FSOUND_GetError)
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_Init() As Boolean
On Local Error GoTo ErrHandler
Dim b As Boolean
ChDir App.Path
b = FSOUND_Init(44100, 32, 0)
If b Then
    FSOUND_DSP_SetActive FSOUND_DSP_GetFFTUnit, 1
    lInitialized = True
    FS_Init = True
Else
    lInitialized = False
    FS_Init = False
End If
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FS_Stream_Open(lFileName As String) As Boolean
On Local Error GoTo ErrHandler
If Len(lFileName) <> 0 Then
    Select Case LCase(Right(lFileName, 4))
    Case ".wav"
    Case ".mp3"
    Case ".ogg"
    Case ".wma"
    Case Else
        Exit Function
    End Select
    lStreamHandle = FSOUND_Stream_Open(lFileName, FSOUND_NORMAL, 0, 0)
    If lStreamHandle <> 0 Then
        FS_Stream_Open = True
    Else
        FS_Stream_Open = False
    End If
Else
    FS_Stream_Open = False
End If
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function FM_PlaySong() As Boolean
On Local Error GoTo ErrHandler
Dim b As Boolean
b = FMUSIC_PlaySong(lSongHandle)
If b Then
    FM_PlaySong = True
Else
    FM_PlaySong = False
End If
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Sub FM_StopSong()
On Local Error GoTo ErrHandler
FMUSIC_StopSong lSongHandle
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Public Sub FS_ClosePlayer()
On Local Error GoTo ErrHandler
FSOUND_Close
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Public Function FM_LoadSong(lFileName As String) As Boolean
On Local Error GoTo ErrHandler
If Len(lFileName) <> 0 Then
    Select Case LCase(Right(lFileName, 3))
    Case "s3m"
    Case ".it"
    Case ".xm"
    Case "mod"
    Case Else
        Exit Function
    End Select
    lSongHandle = FMUSIC_LoadSong(lFileName)
    If lSongHandle <> 0 Then
        FM_LoadSong = True
    Else
        FM_LoadSong = False
    End If
End If
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

