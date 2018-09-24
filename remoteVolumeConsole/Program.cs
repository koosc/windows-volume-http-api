using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Management;
using System.Linq;
using System.Drawing;
using Nancy;
using Nancy.Extensions;
using Nancy.Hosting.Self;
using Newtonsoft.Json;
using System.IO;
using System.Diagnostics;


namespace SetAppVolumne
{

    public class Application
    {
        public string name { get; set; }
        public int volume { get; set; }
        public string icon { get; set; }
        public int pid{ get; set; }
    }


    public class volumeApiModule : NancyModule
    {
        public volumeApiModule()
        {
            //get("/", args => this.response.astext("hello world"));
            Get["/"] = parameters => "OK";
            Get["/getApplications"] = parameters =>
            {
                List<Application> applications;
                applications = EnumerateApplications();
                //float? masterVolume = null;
                //// Get master volume. Applications are based on this as the max level
                //foreach (var application in applications)
                //{
                //    if (application.name.ToLower().Contains("audiosrv.dll")) {
                //        masterVolume = application.volume;
                //    }
                //}

                //foreach (var application in applications)
                //{
                //    if (application.name.ToLower().Contains("audiosrv.dll"))
                //    {
                //        masterVolume = application.volume;
                //    }

                //    if (!application.name.ToLower().Contains("audiosrv.dll"))
                //    {
                //        application.volume = (application.volume / 100) * (masterVolume / 100);
                //        application.volume = application.volume * 100;
                //    }
                //}

                    return JsonConvert.SerializeObject(applications);
            };
            Post["/setVolume"] = parameters =>
            {
                Console.WriteLine("Got set volume request");

                var dataString = this.Request.Body.AsString();

                var data = JsonConvert.DeserializeObject<dynamic>(dataString);

                var requestPid = (int)data.application;
                var requestVolume = (int)data.volume;

                List<Application> applications;
                applications = EnumerateApplications();

                foreach (var application in applications)
                {
                    if (application.pid == requestPid)
                    {
                        Console.WriteLine("Set volume of this application " + requestPid.ToString());
                        SetApplicationVolume(application.pid, (float)requestVolume);
                    }
                }
                //float? masterVolume = null;
                //// Get master volume. Applications are based on this as the max level
                //foreach (var application in applications)
                //{
                //    if (application.name.ToLower().Contains("audiosrv.dll")) {
                //        masterVolume = application.volume;
                //    }
                //}

                //foreach (var application in applications)
                //{
                //    if (application.name.ToLower().Contains("audiosrv.dll"))
                //    {
                //        masterVolume = application.volume;
                //    }

                //    if (!application.name.ToLower().Contains("audiosrv.dll"))
                //    {
                //        application.volume = (application.volume / 100) * (masterVolume / 100);
                //        application.volume = application.volume * 100;
                //    }
                //}

                return JsonConvert.SerializeObject(applications);
            };
        }

        public static float GetApplicationVolume(string name)
        {
            ISimpleAudioVolume volume = GetVolumeObject(name);
            if (volume == null)
                return 0;

            float level;
            volume.GetMasterVolume(out level);
            return level * 100;
        }

        public static bool? GetApplicationMute(string name)
        {
            ISimpleAudioVolume volume = GetVolumeObject(name);
            if (volume == null)
                return null;

            bool mute;
            volume.GetMute(out mute);
            return mute;
        }

        public static void SetApplicationVolume(string name, float level)
        {
            ISimpleAudioVolume volume = GetVolumeObject(name);
            if (volume == null)
                return;

            Guid guid = Guid.Empty;
            volume.SetMasterVolume(level / 100, ref guid);
        }

        public static void SetApplicationVolume(int pid, float level)
        {
            ISimpleAudioVolume volume = GetVolumeObject(pid);
            if (volume == null)
                return;

            Guid guid = Guid.Empty;
            volume.SetMasterVolume(level / 100, ref guid);
        }

        public static void SetApplicationMute(string name, bool mute)
        {
            ISimpleAudioVolume volume = GetVolumeObject(name);
            if (volume == null)
                return;

            Guid guid = Guid.Empty;
            volume.SetMute(mute, ref guid);
        }

        public static List<Application> EnumerateApplications()
        {
            // get the speakers (1st render + multimedia) device

            List<Application> applications = new List<Application>();

            IMMDeviceEnumerator deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
            IMMDevice speakers = null;
            deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out speakers);

            // activate the session manager. we need the enumerator
            Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
            object o;
            speakers.Activate(ref IID_IAudioSessionManager2, 0, IntPtr.Zero, out o);
            IAudioSessionManager2 mgr = (IAudioSessionManager2)o;

            // enumerate sessions for on this device
            IAudioSessionEnumerator sessionEnumerator;
            mgr.GetSessionEnumerator(out sessionEnumerator);
            int count;
            sessionEnumerator.GetCount(out count);

            for (int i = 0; i < count; ++i)
            {
                IAudioSessionControl2 ctl = null;
                sessionEnumerator.GetSession(i, out ctl);

                string dn;
                string iconPath;
                int pid;
                ctl.GetProcessId(out pid);
                ctl.GetDisplayName(out dn);
                float volumeLevel = GetApplicationVolume(dn);
                int roundedVolume = (int)Math.Round(volumeLevel);
                Console.WriteLine("Application Name: " + dn );

                Console.WriteLine(pid);

                Application application = new Application();

                application.name = dn;
                application.volume = roundedVolume;
                application.pid = pid;

                //var proc = Process.GetProcessById(pid);
                //try
                //{
                //    Console.WriteLine(proc.MainModule.FileName);
                //}
                //catch (Exception e)
                //{
                //    Console.WriteLine(e.Message);
                //}
                

                var query = "SELECT ProcessId, Name, ExecutablePath FROM Win32_Process WHERE processid = " + pid.ToString();
                

                using (var searcher = new ManagementObjectSearcher(query))
                using (var results = searcher.Get())
                {
                    var processes = results.Cast<ManagementObject>().Select(x => new
                    {
                        ProcessId = (UInt32)x["ProcessId"],
                        Name = (string)x["Name"],
                        ExecutablePath = (string)x["ExecutablePath"]
                    });
                    Console.WriteLine("Applicaiton name from search: " + processes.ElementAt(0).Name);
                    Console.WriteLine("Applicaiton name from search: " + Process.GetProcessById((int)processes.ElementAt(0).ProcessId).MainWindowTitle);
                    if (application.pid == 0)
                    {
                        application.name = "System Sounds";
                    }
                    else
                    {
                        string mainWindowTitle = Process.GetProcessById((int)processes.ElementAt(0).ProcessId).MainWindowTitle;
                        string executableName = processes.ElementAt(0).Name;
                        if (mainWindowTitle != "")
                        {
                            application.name = mainWindowTitle;
                        }
                        else
                        {
                            application.name = executableName.Split('.')[0];
                        }
                    }
                    



                    foreach (var p in processes)
                    {
                        if (System.IO.File.Exists(p.ExecutablePath))
                        {
                            var icon = Icon.ExtractAssociatedIcon(p.ExecutablePath);

                            icon.ToBitmap().Save("icon.bmp");
                            Byte[] bytes = File.ReadAllBytes("icon.bmp");
                            String encodedIcon = Convert.ToBase64String(bytes);

                            application.icon = encodedIcon;
                            var key = p.ProcessId.ToString();

                            Console.WriteLine(p.ExecutablePath);
                        }
                    }
                }






                applications.Add(application);

                ctl.GetIconPath(out iconPath);
                Console.WriteLine("IconPath: " + iconPath);

                Marshal.ReleaseComObject(ctl);
            }

            return applications;
            Marshal.ReleaseComObject(sessionEnumerator);
            Marshal.ReleaseComObject(mgr);
            Marshal.ReleaseComObject(speakers);
            Marshal.ReleaseComObject(deviceEnumerator);
        }

        private static ISimpleAudioVolume GetVolumeObject(string name)
        {
            // get the speakers (1st render + multimedia) device
            IMMDeviceEnumerator deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
            IMMDevice speakers;
            deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out speakers);

            // activate the session manager. we need the enumerator
            Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
            object o;
            speakers.Activate(ref IID_IAudioSessionManager2, 0, IntPtr.Zero, out o);
            IAudioSessionManager2 mgr = (IAudioSessionManager2)o;

            // enumerate sessions for on this device
            IAudioSessionEnumerator sessionEnumerator;
            mgr.GetSessionEnumerator(out sessionEnumerator);
            int count;
            sessionEnumerator.GetCount(out count);

            // search for an audio session with the required name
            // NOTE: we could also use the process id instead of the app name (with IAudioSessionControl2)
            ISimpleAudioVolume volumeControl = null;
            for (int i = 0; i < count; i++)
            {
                IAudioSessionControl2 ctl;
                sessionEnumerator.GetSession(i, out ctl);
                string dn;
                ctl.GetDisplayName(out dn);
                int pid = ctl.GetProcessId(out pid);
                if (string.Compare(name, dn, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    volumeControl = ctl as ISimpleAudioVolume;
                    break;
                }
                Marshal.ReleaseComObject(ctl);
            }
            Marshal.ReleaseComObject(sessionEnumerator);
            Marshal.ReleaseComObject(mgr);
            Marshal.ReleaseComObject(speakers);
            Marshal.ReleaseComObject(deviceEnumerator);
            return volumeControl;
        }
        private static ISimpleAudioVolume GetVolumeObject(int requestPid)
        {
            // get the speakers (1st render + multimedia) device
            IMMDeviceEnumerator deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
            IMMDevice speakers;
            deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out speakers);

            // activate the session manager. we need the enumerator
            Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
            object o;
            speakers.Activate(ref IID_IAudioSessionManager2, 0, IntPtr.Zero, out o);
            IAudioSessionManager2 mgr = (IAudioSessionManager2)o;

            // enumerate sessions for on this device
            IAudioSessionEnumerator sessionEnumerator;
            mgr.GetSessionEnumerator(out sessionEnumerator);
            int count;
            sessionEnumerator.GetCount(out count);

            // search for an audio session with the required name
            // NOTE: we could also use the process id instead of the app name (with IAudioSessionControl2)
            ISimpleAudioVolume volumeControl = null;
            for (int i = 0; i < count; i++)
            {
                IAudioSessionControl2 ctl;
                sessionEnumerator.GetSession(i, out ctl);
                string dn;
                ctl.GetDisplayName(out dn);
                int pid;
                ctl.GetProcessId(out pid);
                Console.WriteLine("Current pid in search: " + pid.ToString());
                if (pid == requestPid)
                {
                    Console.WriteLine("Matching pid found.");
                    volumeControl = ctl as ISimpleAudioVolume;
                    break;
                }
                Marshal.ReleaseComObject(ctl);
            }
            Marshal.ReleaseComObject(sessionEnumerator);
            Marshal.ReleaseComObject(mgr);
            Marshal.ReleaseComObject(speakers);
            Marshal.ReleaseComObject(deviceEnumerator);
            return volumeControl;
        }
    }

    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    internal class MMDeviceEnumerator
    {
    }

    internal enum EDataFlow
    {
        eRender,
        eCapture,
        eAll,
        EDataFlow_enum_count
    }

    internal enum ERole
    {
        eConsole,
        eMultimedia,
        eCommunications,
        ERole_enum_count
    }

    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceEnumerator
    {
        int NotImpl1();

        [PreserveSig]
        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);

        // the rest is not implemented
    }

    [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDevice
    {
        [PreserveSig]
        int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);

        // the rest is not implemented
    }

    [Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionManager2
    {
        int NotImpl1();
        int NotImpl2();

        [PreserveSig]
        int GetSessionEnumerator(out IAudioSessionEnumerator SessionEnum);

        // the rest is not implemented
    }

    [Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionEnumerator
    {
        [PreserveSig]
        int GetCount(out int SessionCount);

        [PreserveSig]
        int GetSession(int SessionCount, out IAudioSessionControl2 Session);
    }

    [Guid("F4B1A599-7266-4319-A8CA-E70ACB11E8CD"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionControl
    {
        int NotImpl1();

        [PreserveSig]
        int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

        [PreserveSig]
        int SetDisplayName(
            [In] [MarshalAs(UnmanagedType.LPWStr)] string displayName,
            [In] [MarshalAs(UnmanagedType.LPStruct)] Guid eventContext);

        [PreserveSig]
        int GetIconPath([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

        // the rest is not implemented
    }

    [Guid("bfb7ff88-7239-4fc9-8fa2-07c950be9c6d"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioSessionControl2
    {
        // IAudioSessionControl
        [PreserveSig]
        int NotImpl0();

        [PreserveSig]
        int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

        [PreserveSig]
        int SetDisplayName([MarshalAs(UnmanagedType.LPWStr)]string Value, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

        [PreserveSig]
        int GetIconPath([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

        [PreserveSig]
        int SetIconPath([MarshalAs(UnmanagedType.LPWStr)] string Value, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

        [PreserveSig]
        int GetGroupingParam(out Guid pRetVal);

        [PreserveSig]
        int SetGroupingParam([MarshalAs(UnmanagedType.LPStruct)] Guid Override, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

        [PreserveSig]
        int NotImpl1();

        [PreserveSig]
        int NotImpl2();

        // IAudioSessionControl2
        [PreserveSig]
        int GetSessionIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

        [PreserveSig]
        int GetSessionInstanceIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

        [PreserveSig]
        int GetProcessId(out int pRetVal);

        [PreserveSig]
        int IsSystemSoundsSession();

        [PreserveSig]
        int SetDuckingPreference(bool optOut);
    }


    [Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ISimpleAudioVolume
    {
        [PreserveSig]
        int SetMasterVolume(float fLevel, ref Guid EventContext);

        [PreserveSig]
        int GetMasterVolume(out float pfLevel);

        [PreserveSig]
        int SetMute(bool bMute, ref Guid EventContext);

        [PreserveSig]
        int GetMute(out bool pbMute);
    }

    
}


    class Program
    {

        public class hellomodule : NancyModule
        {
            public hellomodule()
            {
                //get("/", args => this.response.astext("hello world"));
                Get["/testing"] = parameters =>
                {
                    return "hello world";
                };
            }
        }

        static void Main(string[] args)
        {
            const string app = "Microsoft Edge";
            HostConfiguration hostConfigs = new HostConfiguration();
            hostConfigs.UrlReservations.CreateAutomatically = true;
            var host = new NancyHost(hostConfigs, new Uri("http://localhost:12345"));
            host.Start();
            Console.ReadKey();

            //foreach (string name in EnumerateApplications())
            //{
            //    Console.WriteLine("name:" + name);
            //    if (name == app)
            //    {
            //        // display mute state & volume level (% of master)
            //        Console.WriteLine("Mute:" + GetApplicationMute(app));
            //        Console.WriteLine("Volume:" + GetApplicationVolume(app));

            //        // mute the application
            //        SetApplicationMute(app, true);

            //        // set the volume to half of master volume (50%)
            //        SetApplicationVolume(app, 50);
            //    }
            //}
        }

        
};