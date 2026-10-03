using System;
using System.Runtime.InteropServices;

namespace SeewoOpt.Services
{
    /// <summary>
    /// 系统音量控制（基于 Core Audio COM API）。
    ///
    /// 整套 COM 互操作定义原本散在 TimeSyncForm 里约 60 行，
    /// 与界面逻辑混在一起，现整体迁入此文件。
    /// </summary>
    public static class VolumeController
    {
        [ComImport]
        [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        private class MMDeviceEnumerator { }

        private enum EDataFlow
        {
            eRender,
            eCapture,
            eAll,
            EDataFlow_enum_count
        }

        private enum ERole
        {
            eConsole,
            eMultimedia,
            eCommunications,
            ERole_enum_count
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
        private interface IMMDeviceEnumerator
        {
            int NotImpl1();

            [PreserveSig]
            int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
        private interface IMMDevice
        {
            [PreserveSig]
            int Activate([MarshalAs(UnmanagedType.LPStruct)] Guid iid, int dwClsCtx,
                IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
        private interface IAudioEndpointVolume
        {
            [PreserveSig] int RegisterControlChangeNotify(IntPtr pNotify);
            [PreserveSig] int UnregisterControlChangeNotify(IntPtr pNotify);
            [PreserveSig] int GetChannelCount(out int pnChannelCount);
            [PreserveSig] int SetMasterVolumeLevel(float fLevelDB, Guid pguidEventContext);
            [PreserveSig] int SetMasterVolumeLevelScalar(float fLevel, Guid pguidEventContext);
            [PreserveSig] int GetMasterVolumeLevel(out float pfLevelDB);
            [PreserveSig] int GetMasterVolumeLevelScalar(out float pfLevel);
            [PreserveSig] int SetChannelVolumeLevel(uint nChannel, float fLevelDB, Guid pguidEventContext);
            [PreserveSig] int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, Guid pguidEventContext);
            [PreserveSig] int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);
            [PreserveSig] int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);
            [PreserveSig] int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, Guid pguidEventContext);
            [PreserveSig] int GetMute(out bool pbMute);
            [PreserveSig] int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount);
            [PreserveSig] int VolumeStepUp(Guid pguidEventContext);
            [PreserveSig] int VolumeStepDown(Guid pguidEventContext);
            [PreserveSig] int QueryHardwareSupport(out uint pdwHardwareSupportMask);
            [PreserveSig] int GetVolumeRange(out float pflVolumeMindB, out float pflVolumeMaxdB, out float pflVolumeIncrementdB);
        }

        private static readonly Guid IID_IAudioEndpointVolume =
            new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");

        /// <summary>
        /// 设置主音量。volumeLevel 取值 0.0 ~ 1.0。
        /// 失败一律返回 false，不抛异常（音量调节失败不应中断校时主流程）。
        /// </summary>
        public static bool SetMasterVolume(float volumeLevel)
        {
            IMMDeviceEnumerator enumerator = null;
            IMMDevice device = null;
            IAudioEndpointVolume volume = null;

            try
            {
                if (volumeLevel < 0.0f || volumeLevel > 1.0f)
                    return false;

                enumerator = new MMDeviceEnumerator() as IMMDeviceEnumerator;
                if (enumerator == null) return false;

                int hr = enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out device);
                if (hr != 0 || device == null) return false;

                hr = device.Activate(IID_IAudioEndpointVolume, 0, IntPtr.Zero, out object obj);
                if (hr != 0 || obj == null) return false;

                volume = obj as IAudioEndpointVolume;
                if (volume == null) return false;

                return volume.SetMasterVolumeLevelScalar(volumeLevel, Guid.Empty) == 0;
            }
            catch
            {
                // 无音频设备、权限不足、服务未启动等情况都归为失败
                return false;
            }
            finally
            {
                // COM 对象必须显式释放，否则音频设备被独占
                if (volume != null) Marshal.ReleaseComObject(volume);
                if (device != null) Marshal.ReleaseComObject(device);
                if (enumerator != null) Marshal.ReleaseComObject(enumerator);
            }
        }

        /// <summary>按百分比设置音量（0-100），内部转换为 0.0~1.0</summary>
        public static bool SetMasterVolumePercent(int percent)
        {
            if (percent < 0 || percent > 100)
                return false;

            return SetMasterVolume(percent / 100.0f);
        }
    }
}
