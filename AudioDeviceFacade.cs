using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

using EasyHook;

using log4net;

using Qisope.PInvoke;
using Qisope.Settings;

namespace Qisope.ServiceProviders.OS
{
    /// <summary>
    /// This class hooks the WinMM (ugh!) windows API methods that Flash uses for audio and replaces some of the functionality
    /// We do this so we can send the Flash audio output to a selected device, rather than the defaul device.
    /// </summary>
    public class AudioDeviceFacade
    {
        private const string WINMM_MODULE_NAME = "winmm.dll";
        private static AudioDeviceFacade s_instance;

        private readonly ILog _log;
        private readonly object _bypassLock = new object();
        private readonly List<int> _bypassThreadIDs = new List<int>();
        private UserSettings _referenceUserSettings;
        private LocalHook _waveOutCloseHook;
        private LocalHook _waveOutGetDevCapsHook;
        private LocalHook _waveOutGetNumDevsHook;
        private LocalHook _waveOutOpenHook;

        public static AudioDeviceFacade Instance
        {
            get
            {
                if (s_instance == null)
                {
                    s_instance = new AudioDeviceFacade();
                }

                return s_instance;
            }
        }

        public AudioDeviceFacade()
        {
            _log = LogManager.GetLogger(GetType());
        }

        ~AudioDeviceFacade()
        {
            UnHook();
        }

        public UserSettings ReferenceUserSettings
        {
            get { return _referenceUserSettings ?? UserSettings.Default; }
            set { _referenceUserSettings = value; }
        }

        public bool BypassLocalOverrides
        {
            get
            {
                lock (_bypassLock)
                {
                    int threadID = Thread.CurrentThread.ManagedThreadId;
                    return !Application.MessageLoop || _bypassThreadIDs.Contains(threadID);
                }
            }
            set
            {
                lock (_bypassLock)
                {
                    int threadID = Thread.CurrentThread.ManagedThreadId;

                    if (value && !_bypassThreadIDs.Contains(threadID))
                    {
                        _bypassThreadIDs.Add(threadID);
                    }
                    else if (!value)
                    {
                        _bypassThreadIDs.Remove(threadID);
                    }
                }
            }
        }

        public void Initialize()
        {
            InstallWinMMHooks();
        }

        public WinMM.tagWAVEOUTCAPSW[] GetWinMMAudioOutputDevices()
        {
            try
            {
                BypassLocalOverrides = true;

                var devices = new List<WinMM.tagWAVEOUTCAPSW>();
                uint devs = WinMM.waveOutGetNumDevs();

                for (uint i = 0; i < devs; i++)
                {
                    var caps = new WinMM.tagWAVEOUTCAPSW();
                    WinMM.waveOutGetDevCaps(i, ref caps, (uint)Marshal.SizeOf(caps));
                    devices.Add(caps);
                }

                return devices.ToArray();
            }
            finally
            {
                BypassLocalOverrides = false;
            }
        }

        private WinMM.WaveOutGetNumDevsDelegate _waveOutGetNumDevsDelegate;
        private WinMM.WaveOutGetDevCapsDelegate _waveOutGetDevCapsDelegate;
        private WinMM.WaveOutOpenDelegate _waveOutOpenDelegate;
        private WinMM.WaveOutCloseDelegate _waveOutCloseDelegate;

        private void InstallWinMMHooks()
        {
            _log.Debug("Setting up Flash audio device facade");

            Kernel32.LoadLibrary(WINMM_MODULE_NAME);

            IntPtr waveOutGetNumDevsPtr = LocalHook.GetProcAddress(WINMM_MODULE_NAME, "waveOutGetNumDevs");
            _waveOutGetNumDevsDelegate = new WinMM.WaveOutGetNumDevsDelegate(OnWaveOutGetNumDevs);
            _waveOutGetNumDevsHook = LocalHook.Create(waveOutGetNumDevsPtr, _waveOutGetNumDevsDelegate, this);
            _waveOutGetNumDevsHook.ThreadACL.SetExclusiveACL(new int[0]);

            IntPtr waveOutGetDevCapsPtr = LocalHook.GetProcAddress(WINMM_MODULE_NAME, "waveOutGetDevCapsA");
            _waveOutGetDevCapsDelegate = new WinMM.WaveOutGetDevCapsDelegate(OnWaveOutGetDevCaps);
            _waveOutGetDevCapsHook = LocalHook.Create(waveOutGetDevCapsPtr, _waveOutGetDevCapsDelegate, this);
            _waveOutGetDevCapsHook.ThreadACL.SetExclusiveACL(new int[0]);

            IntPtr waveOutOpenPtr = LocalHook.GetProcAddress(WINMM_MODULE_NAME, "waveOutOpen");
            _waveOutOpenDelegate = new WinMM.WaveOutOpenDelegate(OnWaveOutOpen);
            _waveOutOpenHook = LocalHook.Create(waveOutOpenPtr, _waveOutOpenDelegate, this);
            _waveOutOpenHook.ThreadACL.SetExclusiveACL(new int[0]);

            IntPtr waveOutClosePtr = LocalHook.GetProcAddress(WINMM_MODULE_NAME, "waveOutClose");
            _waveOutCloseDelegate = new WinMM.WaveOutCloseDelegate(OnWaveOutClose);
            _waveOutCloseHook = LocalHook.Create(waveOutClosePtr, _waveOutCloseDelegate, this);
            _waveOutCloseHook.ThreadACL.SetExclusiveACL(new int[0]);

            BypassLocalOverrides = false;
        }

        public void UnHook()
        {
            GC.SuppressFinalize(this);

            if (_waveOutGetNumDevsHook != null)
            {
                _waveOutGetNumDevsHook.Dispose();
                _waveOutGetNumDevsHook = null;
            }

            if (_waveOutGetDevCapsHook != null)
            {
                _waveOutGetDevCapsHook.Dispose();
                _waveOutGetDevCapsHook = null;
            }

            if (_waveOutOpenHook != null)
            {
                _waveOutOpenHook.Dispose();
                _waveOutOpenHook = null;
            }

            if (_waveOutCloseHook != null)
            {
                _waveOutCloseHook.Dispose();
                _waveOutCloseHook = null;
            }
        }

        private uint OnWaveOutGetNumDevs()
        {
            if (BypassLocalOverrides)
            {
                return WinMM.waveOutGetNumDevs();
            }

            _log.Debug("WaveOutGetNumDevs: Returning 1 device");
            return 1;
        }

        private uint OnWaveOutGetDevCaps(uint udeviceid, ref WinMM.tagWAVEOUTCAPSW pwoc, uint cbwoc)
        {
            if (BypassLocalOverrides)
            {
                return WinMM.waveOutGetDevCaps(udeviceid, ref pwoc, cbwoc);
            }

            _log.Debug("WaveOutGetDevCaps");
            uint deviceID = GetDeviceID();

            uint result = WinMM.waveOutGetDevCaps(deviceID, ref pwoc, cbwoc);

            if (result != 0 && deviceID != WinMM.WAVE_MAPPER)
            {
                result = WinMM.waveOutGetDevCaps(WinMM.WAVE_MAPPER, ref pwoc, cbwoc);
            }

            return result;
        }

        private uint OnWaveOutOpen(ref IntPtr phwo, uint udeviceid, ref WinMM.tWAVEFORMATEX pwfx, uint dwcallback, uint dwcallbackinstance, uint fdwopen)
        {
            if (BypassLocalOverrides)
            {
                return WinMM.waveOutOpen(ref phwo, udeviceid, ref pwfx, dwcallback, dwcallbackinstance, fdwopen);
            }

            _log.Debug("WaveOutOpen");
            uint deviceID = GetDeviceID();

            uint result = WinMM.waveOutOpen(ref phwo, deviceID, ref pwfx, dwcallback, dwcallbackinstance, fdwopen);

            if (result != 0 && deviceID != WinMM.WAVE_MAPPER)
            {
                result = WinMM.waveOutOpen(ref phwo, WinMM.WAVE_MAPPER, ref pwfx, dwcallback, dwcallbackinstance, fdwopen);
            }

            return result;
        }

        private static uint OnWaveOutClose(IntPtr hwo)
        {
            return WinMM.waveOutClose(hwo);
        }

        private uint GetDeviceID()
        {
            _log.Debug("GetDeviceID");

            uint deviceID = WinMM.WAVE_MAPPER;
            WinMM.tagWAVEOUTCAPSW[] devices = GetWinMMAudioOutputDevices();

            _log.DebugFormat("Found {0} devices", devices == null ? 0 : devices.Length);

            if (devices != null && devices.Length > 0)
            {
                int settingsDeviceID = ReferenceUserSettings.AudioOutputDeviceID;
                string settingsDeviceName = ReferenceUserSettings.AudioOutputDevice;

                if (!string.IsNullOrEmpty(settingsDeviceName))
                {
                    _log.DebugFormat("Preferred device: {0}", settingsDeviceName);

                    if (settingsDeviceID >= 0 && devices.Length > settingsDeviceID)
                    {
                        // Try get device by device ID
                        WinMM.tagWAVEOUTCAPSW device = devices[settingsDeviceID];

                        if (settingsDeviceName.StartsWith(device.szPname))
                        {
                            deviceID = (uint)settingsDeviceID;
                        }
                    }

                    if (deviceID == WinMM.WAVE_MAPPER)
                    {
                        // Find device by name.  WAVEOUTCAPSW device names are only partial names - 32 characters - so we do a StartsWith. WinMM sucks.
                        int index = Array.FindIndex(devices, waveoutcapsw => settingsDeviceName.StartsWith(waveoutcapsw.szPname));

                        if (index >= 0)
                        {
                            deviceID = (uint)index;
                        }
                    }
                }
            }

            _log.DebugFormat("Returning device ID: {0}", deviceID == WinMM.WAVE_MAPPER ? "WAVE_MAPPER" : deviceID.ToString());

            return deviceID;
        }

        public class LocalOverrideBypass : IDisposable
        {
            private AudioDeviceFacade _instance;

            public LocalOverrideBypass()
            {
                if (!Instance.BypassLocalOverrides)
                {
                    _instance = Instance;
                    _instance.BypassLocalOverrides = true;
                }
            }

            public void Dispose()
            {
                if (_instance != null)
                {
                    _instance.BypassLocalOverrides = false;
                    _instance = null;
                }
            }
        }
    }
}